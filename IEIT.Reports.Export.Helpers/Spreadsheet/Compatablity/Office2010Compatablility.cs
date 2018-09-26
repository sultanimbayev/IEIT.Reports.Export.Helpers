using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using excel2013 = DocumentFormat.OpenXml.Office2013.Excel;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    /// <summary>
    /// Настройка совместимости документа с Office 2010
    /// </summary>
    public static class Office2010Compatablility
    {
        /// <summary>
        /// Настраивает документ на совместимость с Office 2010
        /// (удаляет элементы которые несовместимы с Office 2010)
        /// </summary>
        /// <param name="doc"></param>
        public static void Office2010Compatablity(this SpreadsheetDocument doc)
        {
            RemoveOffice2013TimelineStyles(doc);
            RemoveUnknownElementsFromWorkbook(doc);
        }

        /// <summary>
        /// Удаление неизвестных элементов
        /// </summary>
        /// <param name="doc"></param>
        public static void RemoveUnknownElementsFromWorkbook(SpreadsheetDocument doc)
        {
            var unknownElements = doc.WorkbookPart.Workbook.Descendants<OpenXmlUnknownElement>();
            _RemoveElements.RemoveElements(unknownElements, deleteSectionIfEmpty: true);
        }

        /// <summary>
        /// Удаленин элементов TimelineStyles
        /// </summary>
        /// <param name="doc"></param>
        public static void RemoveOffice2013TimelineStyles(SpreadsheetDocument doc)
        {
            var stylesheet = doc.GetStylesheet();
            var timelineStyles = stylesheet.Descendants<excel2013.TimelineStyles>();
            _RemoveElements.RemoveElements(timelineStyles, deleteSectionIfEmpty: true);
        }
        
    }
}
