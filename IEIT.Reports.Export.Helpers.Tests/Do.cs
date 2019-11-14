using DocumentFormat.OpenXml.Packaging;
using IEIT.Reports.Export.Helpers.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Tests
{
    public class Do
    {
        public static void ExcelOpen(Action<SpreadsheetDocument> run, bool openFolder = false)
        {
            var guid = Guid.NewGuid();
            var tempDir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), guid.ToString()));
            var filename = $"TestCase-{guid}.xlsx";
            if (!filename.EndsWith(".xlsx")) { filename = filename + ".xlsx"; }
            var filepath = Path.Combine(tempDir.FullName, filename);
            using (var doc = Document.CreateBlank(filepath))
            {
                run(doc);
            }
            var proc = Directory.EnumerateFiles(tempDir.FullName).Count() == 1 && ! openFolder ? Process.Start(filepath) : Process.Start(tempDir.FullName);
            proc.WaitForInputIdle();
        }
    }
}
