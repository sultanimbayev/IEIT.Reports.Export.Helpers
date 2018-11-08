using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet.Intents
{
    public class MakeStyleIntent
    {
        private Stylesheet Stylesheet;
        private CellFormat CellFormat;

        public MakeStyleIntent(Stylesheet stylesheet)
        {
            this.Stylesheet = stylesheet;
            this.CellFormat = new CellFormat();
        }

        public MakeStyleIntent WithAlignment(Alignment alignment)
        {
            if(alignment == null)
            {
                this.CellFormat.Alignment = null;
                this.CellFormat.ApplyAlignment = false;
                return this;
            }
            this.CellFormat.Alignment = alignment;
            this.CellFormat.ApplyAlignment = true;
            return this;
        }

        public MakeStyleIntent WithNumFormat(NumberingFormat numFormat)
        {
            if (numFormat == null)
            {
                this.CellFormat.NumberFormatId = 0;
                this.CellFormat.ApplyNumberFormat = false;
                return this;
            }
            this.CellFormat.NumberFormatId = this.Stylesheet.NumFormat(numFormat);
            this.CellFormat.ApplyNumberFormat = true;
            return this;
        }

        public MakeStyleIntent WithBorder(Border border)
        {
            if(border == null)
            {
                this.CellFormat.BorderId = 0;
                this.CellFormat.ApplyBorder = false;
                return this;
            }
            this.CellFormat.BorderId = this.Stylesheet.MakeBorder(border);
            this.CellFormat.ApplyBorder = true;
            return this;
        }

        public MakeStyleIntent WithFont(Font font)
        {
            if(font == null)
            {
                this.CellFormat.FontId = 0;
                this.CellFormat.ApplyFont = false;
                return this;
            }
            this.CellFormat.FontId = this.Stylesheet.Font(font);
            this.CellFormat.ApplyFont = true;
            return this;
        }

        public MakeStyleIntent WithFill(Fill fill)
        {
            if(fill == null)
            {
                this.CellFormat.FillId = 0;
                this.CellFormat.ApplyFill = false;
            }
            this.CellFormat.FillId = this.Stylesheet.Fill(fill);
            this.CellFormat.ApplyFill = true;
            return this;
        }

        public MakeStyleIntent WithFill(PatternFill patternFill)
        {
            return WithFill(new Fill() { PatternFill = patternFill });
        }

        public MakeStyleIntent WithFill(string rgbColor, PatternValues patternType = PatternValues.Solid)
        {
            var _rgb = rgbColor.TrimStart('#');
            var patternFill = new PatternFill() { PatternType = PatternValues.Solid };
            patternFill.ForegroundColor = new ForegroundColor() { Rgb = HexBinaryValue.FromString(_rgb) };
            return WithFill(patternFill);
        }

        public uint Build()
        {
            return this.Stylesheet.CellFormat(this.CellFormat);
        }

    }
}
