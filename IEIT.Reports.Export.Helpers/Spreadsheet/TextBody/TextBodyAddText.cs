﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class TextBodyAddText
    {
        public static xdr.TextBody AddText(this xdr.TextBody textBody, string text, Font font = null, Color? fontColor = null)
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
