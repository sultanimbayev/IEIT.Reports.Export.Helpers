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
                var guid = Guid.NewGuid();
                doc.SetCustomTextProperty("myProp", guid.ToString());
                var propValue = doc.GetCustomPropertyAsString("myProp");
                var ws = doc.GetWorksheets().First();
                ws.Write(propValue).To("B2");
            });
        }
        
    }
}
