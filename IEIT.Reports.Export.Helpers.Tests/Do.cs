using DocumentFormat.OpenXml.Packaging;
using IEIT.Reports.Export.Helpers.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Tests
{
    public class Do
    {
        public static Process ExcelOpen(Action<SpreadsheetDocument> run, bool openFolder = false)
        {
            return GenerateFilesIn((tempDir) =>
            {
                var guid = Guid.NewGuid();
                var filename = $"TestCase-{guid}.xlsx";
                if (!filename.EndsWith(".xlsx")) { filename = filename + ".xlsx"; }
                var filepath = Path.Combine(tempDir, filename);
                using (var doc = Document.CreateWorkbook(filepath))
                {
                    run(doc);
                }
            }, openFolder : openFolder);
        }

        public static Process GenerateFilesIn(Action<string> run, bool openFolder = false)
        {
            var guid = Guid.NewGuid();
            var tempDir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), guid.ToString()));
            run(tempDir.FullName);
            var files = Directory.EnumerateFiles(tempDir.FullName).ToList();
            var proc = (files.Count == 1 && !openFolder) ? Process.Start(files.FirstOrDefault()) : Process.Start(tempDir.FullName);
            return proc;
        }

        public static string GetProjectDir()
        {
            string codeBase = Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            string path = Uri.UnescapeDataString(uri.Path);
            var binDir = Path.GetDirectoryName(path);
            var projectDir = Path.GetFullPath(Path.Combine(binDir, "../../"));
            return projectDir;
        }
    }
}
