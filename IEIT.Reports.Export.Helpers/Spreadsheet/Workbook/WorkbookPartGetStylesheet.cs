using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorkbookPartGetStylesheet
    {

        /// <summary>
        /// Получить таблицу со стилями
        /// </summary>
        /// <param name="wbPart">Рабочая книга документа</param>
        /// <returns>Таблицу содержащяя стили документа</returns>
        public static Stylesheet GetStylesheet(this WorkbookPart wbPart)
        {
            if (wbPart == null) { throw new ArgumentNullException("WorkbookPart must be not null!"); }

            if (wbPart.WorkbookStylesPart == null) { wbPart.AddNewPart<WorkbookStylesPart>(); }
            if (wbPart.WorkbookStylesPart.Stylesheet == null) { wbPart.WorkbookStylesPart.Stylesheet = new Stylesheet(); }
            return wbPart.WorkbookStylesPart.Stylesheet;
        }
    }
}
