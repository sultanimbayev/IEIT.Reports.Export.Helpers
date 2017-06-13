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

        internal static SharedStringItem GetSharedStringItem(this WorkbookPart wbPart, int itemId)
        {
            if(wbPart == null) { throw new ArgumentNullException("wbPart is null"); }
            return wbPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(itemId);
        }

        public static SharedStringItem Add(this SharedStringTable sst, string text, RunProperties rPr = null)
        {
            var item = new SharedStringItem();
            var run = new Run();
            run.Text = new Text(text);
            if(rPr != null) { run.Append(rPr.CloneNode(true) as RunProperties); }
            item.Append(run);
            sst.Append(item);
            sst.Count.Value++;
            return item;
        }

        public static void Append(this SharedStringItem item, string text, RunProperties rPr = null)
        {
            if(item == null) { throw new ArgumentNullException($"SharedStringItem object is null"); }
            var lastElem = item.Elements<Run>().LastOrDefault();
            if (!lastElem.RunProperties.SameAs(rPr))
            {
                var run = new Run();
                run.RunProperties = rPr;
                run.Text = new Text(text);
                item.InsertAfter(run, lastElem);
                return;
            }

            if(lastElem.Text == null || string.IsNullOrEmpty(lastElem.Text.Text))
            {
                lastElem.Text = new Text(text);
                return;
            }

            lastElem.Text.Text += text;
            return;
            
        }

    }
}
