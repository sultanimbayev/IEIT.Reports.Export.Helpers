using DocumentFormat.OpenXml;
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
    public class GetWorkbookIdTest
    {
        [TestCase]
        public void DefaultAttibuteWhenNoWorkbookIdIsSet()
        {
            Do.ExcelOpen(doc =>
            {
                var wb = doc.WorkbookPart.Workbook;
                var wbIdAttr = wb.ExtendedAttributes
                    .FirstOrDefault(attr => attr.LocalName == "workbookId" && string.IsNullOrEmpty(attr.NamespaceUri));
                Assert.AreEqual(wbIdAttr, default(OpenXmlAttribute));
            });
        }

        [TestCase]
        public void GetWorkbookIdTestCaseOne()
        {
            Do.ExcelOpen(doc =>
            {
                var wb = doc.WorkbookPart.Workbook;

                var guid = Guid.NewGuid();
                wb.SetAttribute(new OpenXmlAttribute()
                {
                    LocalName = "workbookId",
                    Value = guid.ToString()
                });
                var wbIdAttr = wb.ExtendedAttributes
                    .FirstOrDefault(attr => attr.LocalName == "workbookId" && string.IsNullOrEmpty(attr.NamespaceUri));

                var wbIdActual = wb.GetWorkbookId();
                var wbIdExpected = wbIdAttr.Value;
                Assert.AreEqual(wbIdExpected, wbIdActual); ;
            });
        }

        [TestCase]
        public void GetWorkbookIdTestCaseTwo()
        {
            Do.ExcelOpen(doc =>
            {
                var wb = doc.WorkbookPart.Workbook;

                var guid = Guid.NewGuid();
                var wbIdActual = wb.GetWorkbookId();
                var wbIdAttr = wb.ExtendedAttributes
                    .FirstOrDefault(attr => attr.LocalName == "workbookId" && string.IsNullOrEmpty(attr.NamespaceUri));

                var wbIdExpected = wbIdAttr.Value;
                Assert.AreEqual(wbIdExpected, wbIdActual); ;
            });
        }

        [TestCase]
        public void GetWorkbookIdRemainsTheSame()
        {
            Do.ExcelOpen(doc =>
            {
                var wb = doc.WorkbookPart.Workbook;
                var wbId1 = wb.GetWorkbookId();
                var wbId2 = wb.GetWorkbookId();
                Assert.AreEqual(wbId2, wbId1); ;
            });
        }
    }
}
