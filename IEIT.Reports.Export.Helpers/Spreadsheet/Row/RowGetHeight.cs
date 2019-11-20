using IEIT.Reports.Export.Helpers.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using x = DocumentFormat.OpenXml.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class RowGetHeight
    {
        public static double GetHeight(this x.Row row)
        {
            var height = row.Height;
            if (height != null) { return height.Value; }
            var ws = row.ParentOfType<x.Worksheet>();
            if(ws == null)
            {
                throw new InvalidDocumentStructureException();
            }
            var sheetFormatProps = ws.SheetFormatProperties;
            if(sheetFormatProps == null)
            {
                sheetFormatProps = new x.SheetFormatProperties();
                ws.Insert(sheetFormatProps).AfterOneOf(typeof(x.Dimension), typeof(x.SheetView));
            }
            return sheetFormatProps.DefaultRowHeight ?? 14.4;
        }

        /// <summary>
        /// Height in pixels with 96dpi
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public static double GetHeightInPixels(this x.Row row)
        {
            return row.GetHeight() * 96 / 72;
        }
    }
}
