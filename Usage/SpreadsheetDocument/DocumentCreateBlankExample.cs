using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using IEIT.Reports.Export.Helpers.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Usage.Interfaces;

namespace Usage
{
    public class DocumentCreateBlankExample : ICreateFile
    {
        public string CreateOne()
        {
            var guid = Guid.NewGuid();
            var filepath = guid.ToString() + ".xlsx";
            var doc = Document.CreateBlank(filepath);
            doc.SaveAndClose();
            return filepath;
        }
    }
}
