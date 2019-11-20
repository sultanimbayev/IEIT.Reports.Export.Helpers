using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class DocumentSave
    {
        /// <summary>
        /// Сохранить изменения в документе
        /// </summary>
        /// <param name="document">Документ над которым производится операция</param>
        public static void Save(this SpreadsheetDocument document)
        {
            document.WorkbookPart.Workbook.Save();
        }
    }
}
