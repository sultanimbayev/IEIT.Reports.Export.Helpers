using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;


namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ParagraphAddText
    {
        public static a.Paragraph AddText(this a.Paragraph paragraph, string text, Font font = null, Color? fontColor = null)
        {
            var run = new a.Run().Init(text, font, fontColor);
            paragraph.Append(run);
            return paragraph;
        }
    }
}
