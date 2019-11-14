using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class TextBodySetHorizontalAlignment
    {
        public static a.TextBody SetHorizontalAlignment(this a.TextBody textBody, a.TextAlignmentTypeValues textAlignment)
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
