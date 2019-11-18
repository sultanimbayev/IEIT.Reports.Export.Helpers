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
    public class AddPictureTest
    {
        [TestCase]
        public void TestCase()
        {
            Do.ExcelOpen((doc) =>
            {
                var ws = doc.GetWorksheets().First();
                var projectDir = Do.GetProjectDir();
                var path = Path.Combine(projectDir, "images/happy-bday.jpg");
                var picture = ws.AddPicture(path);
                picture.SetTopLeft("B3");
                picture.SetBottomRight("F16");
                doc.SaveAndClose();
            });
        }
    }
}
