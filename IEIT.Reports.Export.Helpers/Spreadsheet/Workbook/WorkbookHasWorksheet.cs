using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorkbookHasWorksheet
    {
        /// <summary>
        /// Получить информацию о существовании листа с указанным названием
        /// </summary>
        /// <param name="workbook">Рабочая книга документа</param>
        /// <param name="sheetName">Название листа</param>
        /// <returns>true если лист с таким названием существует в книге, false в обратном случае</returns>
        public static bool HasWorksheet(this Workbook workbook, string sheetName)
        {
            if (workbook == null) { throw new ArgumentNullException("workbook"); }
            if (workbook.WorkbookPart == null) { throw new InvalidDocumentStructureException(); }
            return workbook.Descendants<Sheet>()
                .Where(s => s.Name.Value.Equals(sheetName))
                .Count() > 0;
        }
    }
}
