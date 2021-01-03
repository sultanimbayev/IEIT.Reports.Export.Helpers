using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
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
    public class DocumentIdTest
    {
        [TestCase]
        public void TestCase()
        {
            Do.ExcelOpen((doc) =>
            {
                var documentId = doc.GetDocumentId();
                var documentId2 = doc.GetDocumentId();
                Assert.AreEqual(documentId, documentId2);
                var ws = doc.GetWorksheets().First();
                ws.Write("Document id:").To("A2");
                ws.Write("Document id 2:").To("A3");
                ws.Write(documentId).To("B2");
                ws.Write(documentId2).To("B3");
                ws.GetColumn("A").SetWidth(20d);
            });
        }
        
    }
}
