using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class TextBodySetVerticalAlignment
    {
        public static xdr.TextBody SetVerticalAlignment(this xdr.TextBody textBody, a.TextAnchoringTypeValues alignment)
        {
            var textBodyProps = textBody.GetFirstChild<a.BodyProperties>();
            if(textBodyProps == null)
            {
                textBodyProps = new a.BodyProperties()
                {
                    VerticalOverflow = a.TextVerticalOverflowValues.Clip,
                    HorizontalOverflow = a.TextHorizontalOverflowValues.Clip
                };
                textBody.PrependChild(textBodyProps);
            }
            textBodyProps.Anchor = alignment;
            return textBody;
        }
    }
}
