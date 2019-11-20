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
    public class ColumnWidthGetSetTest
    {
        [TestCase]
        public void TestCase1()
        {
            Do.ExcelOpen((doc) =>
            {
                var ws = doc.GetWorksheets().First();
                var columnB = ws.GetColumn("B");
                var oldWidth = columnB.GetWidthInPixels();
                var newWidth = 200;
                columnB.SetWidthInPixels(newWidth);
                var newWidth2 = columnB.GetWidthInPixels();
                //Assert.AreEqual(newWidth, newWidth2);
                ws.Write("<- Initial width").To("C2");
                ws.Write(oldWidth).To("B2");

                ws.Write("<- New width").To("C3");
                ws.Write(newWidth2).To("B3");
            });
        }
    }
}
