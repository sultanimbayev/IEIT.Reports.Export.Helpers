using dr = DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using DocumentFormat.OpenXml;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class TextCharacterPropertiesTypeSetFont
    {
        public static dr.TextCharacterPropertiesType SetFont(this dr.TextCharacterPropertiesType props, Font font, Color? fontColor = null)
        {
            if(props == null) { return props; }
            if(font == null) { font = new Font("Calibri", 11); }
            props.FontSize = (int)(font.Size * 100);
            props.Italic = font.Italic;
            props.Bold = font.Bold;
            props.Underline = font.Underline ? dr.TextUnderlineValues.Dash : dr.TextUnderlineValues.None;
            var fontFill = props.GetFirstChild<dr.SolidFill>();
            if(fontFill == null)
            {
                fontFill = new dr.SolidFill();
                props.PrependChild(fontFill);
                var sysColor = new dr.SystemColor();
                fontFill.Append(sysColor);
                sysColor.Val = dr.SystemColorValues.WindowText;
                fontColor = fontColor ?? Color.Black;
                sysColor.LastColor = new HexBinaryValue(fontColor.Value.ToHex());
            }
            var latinFont = props.GetFirstChild<dr.LatinFont>();
            if(latinFont == null)
            {
                latinFont = new dr.LatinFont();
                props.InsertAfter(latinFont, fontFill);
            }
            var complexScriptFont = props.GetFirstChild<dr.ComplexScriptFont>();
            if(complexScriptFont == null)
            {
                complexScriptFont = new dr.ComplexScriptFont();
                props.InsertAfter(complexScriptFont, latinFont);
            }
            latinFont.Typeface = font.FontFamily.Name;
            latinFont.CharacterSet = 0;
            complexScriptFont.Typeface = font.FontFamily.Name;
            complexScriptFont.CharacterSet = 0;
            return props;
        }
    }
}
