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
    public class RowHeightGetSetTest
    {
        [TestCase]
        public void TestCase1()
        {
            Do.ExcelOpen((doc) =>
            {
                var ws = doc.GetWorksheets().First();
                var row = ws.GetRow(2);
                var initialHeight = row.GetHeightInPixels(dpi:96);
                row.SetHeightInPixels(24, dpi:96);
                var newHeight = row.GetHeightInPixels(dpi:96);
                ws.Write(initialHeight).To("B2");
                ws.Write(newHeight).To("B3");
            });
        }
    }
}
