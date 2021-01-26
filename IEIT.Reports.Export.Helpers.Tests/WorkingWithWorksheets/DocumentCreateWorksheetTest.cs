using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
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
    public class DocumentCreateWorksheetTest
    {
        [TestCase]
        public void CreatingWorksheetOnEmptyDocument()
        {
            Do.GenerateFilesIn((dir) =>
            {
                using (var doc = SpreadsheetDocument.Create(Path.Combine(dir, "test.xlsx"), SpreadsheetDocumentType.Workbook))
                {
                    doc.NewWorksheet("HelloWorld");
                    doc.NewWorksheet("HelloAgain!");
                }
            });
        }
    }
}
