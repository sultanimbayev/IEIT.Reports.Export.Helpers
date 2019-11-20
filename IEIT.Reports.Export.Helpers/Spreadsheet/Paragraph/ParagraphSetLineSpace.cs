using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ParagraphSetLineSpace
    {
        public static a.Paragraph SetLineSpace(this a.Paragraph paragraph, float heightMultiplier)
        {
            var paragraphProperties = paragraph.GetFirstChild<a.ParagraphProperties>();
            if (paragraphProperties == null)
            {
                paragraphProperties = new a.ParagraphProperties();
                paragraph.PrependChild(paragraphProperties);
            }
            paragraphProperties.LineSpacing = new a.LineSpacing(new a.SpacingPercent() { Val = (int)(heightMultiplier * 100000) });
            return paragraph;
        }
    }
}
