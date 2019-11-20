using x = DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IEIT.Reports.Export.Helpers.Exceptions;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class RowSetHeight
    {
        public static x.Row SetHeight(this x.Row row, double height)
        {
            if(row == null) { return null; }
            var ws = row.ParentOfType<x.Worksheet>();
            if (ws == null)
            {
                throw new InvalidDocumentStructureException();
            }
            var sheetFormatProps = ws.SheetFormatProperties;
            if (sheetFormatProps == null)
            {
                sheetFormatProps = new x.SheetFormatProperties();
                ws.Insert(sheetFormatProps).AfterOneOf(typeof(x.Dimension), typeof(x.SheetView));
            }
            if (sheetFormatProps.DefaultRowHeight == null || !sheetFormatProps.DefaultRowHeight.HasValue)
            {
                sheetFormatProps.DefaultRowHeight = 14.4;
            }
            row.Height = height;
            row.CustomHeight = true;
            return row;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="heightInPixels">Pixels in 96dpi</param>
        /// <returns></returns>
        public static x.Row SetHeightInPixels(this x.Row row, double heightInPixels)
        {
            return row.SetHeight(heightInPixels * 72 / 96);
        }
    }
}
