using IEIT.Reports.Export.Helpers.Spreadsheet;
using NUnit.Framework;
using System.Drawing;
using System.Linq;
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
                var shape = ws.AddShape(a.ShapeTypeValues.Rectangle);

                shape.SetTopLeft("B2")
                    .SetBottomRight("E5");

                shape.SetSolidFill(Color.Yellow, 0.5f);
                shape.SetOutlineWidthInPixels(1.5f);
                shape.SetText("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                    "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                    "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris " +
                    "nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in " +
                    "reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla " +
                    "pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa" +
                    " qui officia deserunt mollit anim id est laborum.",
                    new Font("Arial Cyr", 10, FontStyle.Italic), Color.Blue);
                shape.SetLineSpace(0.8f);

                var shape2 = ws.AddShape(a.ShapeTypeValues.Rectangle);
                shape2.SetTopLeft("G2")
                    .SetBottomRight("J5");
                shape2.SetText("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed " +
                    "do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad " +
                    "minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex " +
                    "ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate " +
                    "velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat " +
                    "cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.");
                shape2.SetSolidFill(Color.Red, 0.5f);
                shape2.SetOutlineWidthInPixels(3.5f);
                shape2.SetOutlineColor(Color.Blue);

                doc.SaveAndClose();
            });
        }

    }
}
