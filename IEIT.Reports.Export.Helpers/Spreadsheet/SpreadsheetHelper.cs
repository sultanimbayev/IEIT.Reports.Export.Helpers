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
        /// Получить рабочий лист
        /// </summary>
        /// <param name="doc">Документ, из которого нужно получить лист</param>
        /// <param name="sheetName">Название требуемого листа</param>
        /// <returns>Рабочий лист, название которого соответсвует указанному, или null если лист не найден.</returns>
        public static Worksheet GetWorksheet(this SpreadsheetDocument doc, string sheetName)
        {
            if (doc == null) { throw new ArgumentNullException("doc"); }
            if (doc.WorkbookPart == null || doc.WorkbookPart.Workbook == null) { throw new InvalidDocumentStructureException(); }
            return doc.WorkbookPart.Workbook.GetWorksheet(sheetName);
        }

        /// <summary>
        /// Получить часть документа с рабочим листом
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static Worksheet GetWorksheet(this Workbook workbook, string sheetName)
        {
            if (workbook == null) { throw new ArgumentNullException("workbook"); }
            if (workbook.WorkbookPart == null) { throw new InvalidDocumentStructureException(); }
            var rel = workbook.Descendants<Sheet>()
                .Where(s => s.Name.Value.Equals(sheetName))
                .First();
            if (rel == null || rel.Id == null) { return null; }
            var wsPart = workbook.WorkbookPart.GetPartById(rel.Id) as WorksheetPart;
            if (wsPart == null) { return null; }
            return wsPart.Worksheet;
        }

        /// <summary>
        /// Получить <see cref="SharedStringTable"/> по ID
        /// </summary>
        /// <param name="wbPart">Элемент <see cref="WorkbookPart"/></param>
        /// <param name="itemId">ID элемента <see cref="SharedStringItem"/></param>
        /// <returns>Элемент <see cref="SharedStringItem"/> с указанным ID</returns>
        internal static SharedStringItem GetSharedStringItem(this WorkbookPart wbPart, int itemId)
        {
            if(wbPart == null) { throw new ArgumentNullException("wbPart is null"); }
            return wbPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(itemId);
        }

        /// <summary>
        /// Добавить текст в таблицу <see cref="SharedStringTable"/>
        /// </summary>
        /// <param name="sst">Таблица с тектами</param>
        /// <param name="text">Добавляемый текст</param>
        /// <param name="rPr">Стиль добавляемого текста</param>
        /// <returns>Добавленыый элемент в <see cref="SharedStringTable"/> содержащий указанный текст</returns>
        public static SharedStringItem Add(this SharedStringTable sst, string text, RunProperties rPr = null)
        {
            var item = new SharedStringItem();
            if (rPr == null)
            {
                item.Text = new Text(text);
            }
            else
            {
                var run = new Run();
                run.Text = new Text(text);
                run.Append(rPr.CloneNode(true));
                item.Append(run);
            }
            sst.Append(item);
            if (sst.Count != null) { sst.Count.Value++; }
            if (sst.UniqueCount != null) { sst.UniqueCount.Value++; }
            return item;
        }
        

        /// <summary>
        /// Сохранить изменения и закрыть документ
        /// </summary>
        /// <param name="document">Документ над которым производится операция</param>
        public static void SaveAndClose(this SpreadsheetDocument document)
        {
            document.Save();
            document.Close();
        }
        

        public static Stylesheet GetStylesheet(this SpreadsheetDocument document)
        {
            return document.WorkbookPart.GetStylesheet();
        }

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
