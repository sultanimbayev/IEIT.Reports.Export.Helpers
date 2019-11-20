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
                var initialHeight = row.GetHeight();
                row.SetHeight(60);
                var newHeight = row.GetHeight();
                ws.Write(initialHeight).To("B2");
                ws.Write(newHeight).To("B3");
            }, true);
        }
    }
}
