using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ParagraphSetFont
    {
        public static a.Paragraph SetFont(this a.Paragraph paragraph, Font font, Color? fontColor)
        {
            var runs = paragraph.Descendants<a.Run>();
            var runsCount = 0;
            foreach(var run in runs)
            {
                run.SetFont(font, fontColor);
                runsCount++;
            }
            if(runsCount == 0)
            {
                var r = new a.Run().Init("", font, fontColor);
                paragraph.Append(r);
            }
            return paragraph;
        }
    }
}
