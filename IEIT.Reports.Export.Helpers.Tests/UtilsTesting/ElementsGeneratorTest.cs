using IEIT.Reports.Export.Helpers.Spreadsheet;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Tests.UtilsTesting
{
    [TestFixture]
    public class ElementsGeneratorTest
    {
        [TestCase]
        public void CreatingChartTitleFromScratch()
        {
            var newTitle = new DocumentFormat.OpenXml.Drawing.Charts.Title().From("Drawing\\ChartTitle");
        }
    }
}
