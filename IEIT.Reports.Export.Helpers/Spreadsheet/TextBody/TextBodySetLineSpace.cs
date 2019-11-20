using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using a = DocumentFormat.OpenXml.Drawing;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class TextBodySetLineSpace
    {
        public static xdr.TextBody SetLineSpace(this xdr.TextBody textBody, float heightMultiplier)
        {
            var paragraphs = textBody.Elements<a.Paragraph>();
            foreach (var p in paragraphs)
            {
                p.SetLineSpace(heightMultiplier);
            }
            return textBody;
        }
    }
}
