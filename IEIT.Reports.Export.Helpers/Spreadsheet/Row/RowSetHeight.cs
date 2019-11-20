using x = DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IEIT.Reports.Export.Helpers.Exceptions;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{

    /// <summary>
    /// Row set height extension
    /// </summary>
    public static class RowSetHeight
    {
        /// <summary>
        /// Set height of row in points (pixels with 72dpi)
        /// </summary>
        /// <param name="row">row instance which height is going to change</param>
        /// <param name="height">height of row in points (pixels with 72dpi)</param>
        /// <returns></returns>
        public static x.Row SetHeightInPoints(this x.Row row, double height)
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
                sheetFormatProps.DefaultRowHeight = 18;
            }
            row.Height = height;
            row.CustomHeight = true;
            return row;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="heightInPixels">Pixels</param>
        /// <param name="dpi">Pixels density, dots per inch</param>
        /// <returns></returns>
        public static x.Row SetHeightInPixels(this x.Row row, double heightInPixels, double dpi = 96)
        {
            return row.SetHeightInPoints(heightInPixels / dpi * 72d );
        }
    }
}
