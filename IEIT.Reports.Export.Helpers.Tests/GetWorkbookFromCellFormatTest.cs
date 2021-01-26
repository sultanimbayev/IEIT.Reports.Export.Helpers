using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Spreadsheet;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Tests
{
    [TestFixture]
    public class GetWorkbookFromCellFormatTest
    {
        [TestCase]
        public void GetWorkbookFromCellFormat()
        {
            Do.ExcelOpen(doc =>
            {
                
                var stylesDocPath = @"..\Assets\StylesForTest";
                using (var stylesDoc = SpreadsheetDocument.Open(stylesDocPath, false))
                {
                    var styles = stylesDoc.GetWorksheets().First().GetStyles();
                    var ws = doc.GetWorksheets().First();
                    ws.Write("Hello").To("B2").WithStyle(styles["table_header"]);
                }
            });
        }
    }
}
