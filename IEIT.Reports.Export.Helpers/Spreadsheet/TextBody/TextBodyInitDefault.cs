using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class TextBodyInitDefault
    {
        public static xdr.TextBody InitDefault(this xdr.TextBody textBody)
        {
            if (textBody == null) { return null; }
            var bodyProps = new a.BodyProperties() // BodyProps - user def
            {
                VerticalOverflow = a.TextVerticalOverflowValues.Clip,
                HorizontalOverflow = a.TextHorizontalOverflowValues.Clip,
                RightToLeftColumns = false,
                Anchor = a.TextAnchoringTypeValues.Center
            };
            textBody.Append(bodyProps);

            textBody.Append(new a.ListStyle());

            var paragraph = new a.Paragraph();
            textBody.Append(paragraph); // paragraph - userDef

            return textBody;
        }
    }
}
