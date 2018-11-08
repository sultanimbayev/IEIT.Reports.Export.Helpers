using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetGetStylesheet
    {

        /// <summary>
        /// Получить таблицу стилей
        /// </summary>
        /// <param name="workbook">Рабочая книга документа</param>
        /// <returns>Таблица стилей указанной рабочей книги</returns>
        public static Stylesheet GetStylesheet(this Workbook workbook)
        {
            return workbook.WorkbookPart.GetStylesheet();
        }
    }
}
