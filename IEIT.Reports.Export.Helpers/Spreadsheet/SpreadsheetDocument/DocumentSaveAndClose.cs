using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class DocumentSaveAndClose
    {
        /// <summary>
        /// Сохранить изменения и закрыть документ
        /// </summary>
        /// <param name="document">Документ над которым производится операция</param>
        public static void SaveAndClose(this SpreadsheetDocument document)
        {
            if(document.FileOpenAccess == System.IO.FileAccess.ReadWrite)
            {
                document.WorkbookPart.Workbook.Save();
            }
            document.Close();
        }
    }
}
