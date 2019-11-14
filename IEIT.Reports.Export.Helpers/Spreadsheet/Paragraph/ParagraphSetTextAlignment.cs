using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class ParagraphSetTextAlignment
    {
        public static a.Paragraph SetTextAlignment(this a.Paragraph paragraph, a.TextAlignmentTypeValues textAlignment)
        {
            var paragraphProperties = paragraph.GetFirstChild<a.ParagraphProperties>();
            if(paragraphProperties == null)
            {
                paragraphProperties = new a.ParagraphProperties();
                paragraph.PrependChild(paragraphProperties);
            }
            paragraphProperties.Alignment = textAlignment;
            return paragraph;
        }
    }
}
