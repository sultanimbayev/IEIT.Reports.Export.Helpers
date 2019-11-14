using DocumentFormat.OpenXml.Packaging;
using IEIT.Reports.Export.Helpers.Spreadsheet;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using a = DocumentFormat.OpenXml.Drawing;

namespace IEIT.Reports.Export.Helpers.Tests
{
    [TestFixture]
    public class AddShapeTest
    {
        [TestCase]
        public void TestCase1()
        {
            Do.ExcelOpen((doc) =>
            {
                var ws = doc.GetWorksheets().First();
                var shape = ws.AddShape("rect", a.ShapeTypeValues.Rectangle);
                var shape2 = ws.AddShape("rect", a.ShapeTypeValues.Rectangle);
                doc.SaveAndClose();
            });
        }

    }
}
