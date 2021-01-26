using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class StylesheetGetWorkbookPart
    {
        public static WorkbookPart GetWorkbookPart(this Stylesheet stylesheet)
        {
            return stylesheet.WorkbookStylesPart.GetParentParts().FirstOrDefault() as WorkbookPart;
        }
    }
}
