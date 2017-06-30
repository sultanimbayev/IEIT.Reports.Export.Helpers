using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using IEIT.Reports.Export.Helpers.Exceptions;
using IEIT.Reports.Export.Helpers.Spreadsheet.Intents;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class SpreadsheetHelper
    {
        

        /// <summary>
        /// Сохранить изменения и закрыть документ
        /// </summary>
        /// <param name="document">Документ над которым производится операция</param>
        public static void SaveAndClose(this SpreadsheetDocument document)
        {
            document.Save();
            document.Close();
        }

        /// <summary>
        /// Получить таблицу стилей
        /// </summary>
        /// <param name="document">Рабочий документ</param>
        /// <returns>Таблица стилей указанного документа</returns>
        public static Stylesheet GetStylesheet(this SpreadsheetDocument document)
        {
            return document.WorkbookPart.GetStylesheet();
        }
        
        /// <summary>
        /// Получить таблицу стилей
        /// </summary>
        /// <param name="workbook">Рабочая книга документа</param>
        /// <returns>Таблица стилей указанной рабочей книги</returns>
        public static Stylesheet GetStylesheet(this Workbook workbook)
        {
            return workbook.WorkbookPart.GetStylesheet();
        }

        /// <summary>
        /// Получить таблицу со стилями
        /// </summary>
        /// <param name="wbPart">Рабочая книга документа</param>
        /// <returns>Таблицу содержащяя стили документа</returns>
        public static Stylesheet GetStylesheet(this WorkbookPart wbPart)
        {
            if(wbPart == null) { throw new ArgumentNullException("WorkbookPart must be not null!"); }

            if(wbPart.WorkbookStylesPart == null) { wbPart.AddNewPart<WorkbookStylesPart>(); }
            if(wbPart.WorkbookStylesPart.Stylesheet == null) { wbPart.WorkbookStylesPart.Stylesheet = new Stylesheet(); }
            return wbPart.WorkbookStylesPart.Stylesheet;
        }

        public static void Add(this DifferentialFormats formatsList, DifferentialFormat format)
        {
            formatsList.Append(format);
            if(formatsList.Count != null) { formatsList.Count.Value++; }
        }

    }
}
