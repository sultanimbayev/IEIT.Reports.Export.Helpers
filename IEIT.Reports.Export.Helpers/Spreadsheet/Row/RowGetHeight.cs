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
            var result = sheetFormatProps?.DefaultRowHeight == null ? 18 : sheetFormatProps.DefaultRowHeight.Value * (72d/96d);
            return result;
        }

        /// <summary>
        /// Height in pixels with 96dpi
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public static double GetHeightInPixels(this x.Row row, double dpi = 96)
        {
            return row.GetHeight() * dpi / 72d;
        }
    }
}
