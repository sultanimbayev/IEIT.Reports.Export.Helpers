using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Spreadsheet;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Tests
{
    [TestFixture]
    public class StylesSupportTest
    {
        [TestCase]
        public void GetStylesFromAnotherWorkbook()
        {
            Do.ExcelOpen(doc =>
            {
                var testDir = TestContext.CurrentContext.TestDirectory;
                var stylesDocPath = Path.Combine(testDir, @".\Assets\StylesForTest.xlsx");
                using (var stylesDoc = SpreadsheetDocument.Open(stylesDocPath, false))
                {
                    var styles = stylesDoc.GetWorksheets().First().GetStyles();
                    var ws = doc.GetWorksheets().First();
                    ws.Write("Hello").To("B2").WithStyle(styles["HugeText"]);
                }
            });
        }
    }
}
