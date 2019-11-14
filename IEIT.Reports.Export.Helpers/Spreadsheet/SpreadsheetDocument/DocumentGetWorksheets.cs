using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class DocumentGetWorksheets
    {
        public static IEnumerable<Worksheet> GetWorksheets(this SpreadsheetDocument document)
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>();
            if (sheets.Count() == 0) { return null; }
            var worksheetParts = sheets.Select(s => (WorksheetPart)document.WorkbookPart.GetPartById(s.Id));
            return worksheetParts.Select(wp => wp.Worksheet);
        }
    }
}
