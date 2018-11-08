using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorkbookPartGetSharedStringItem
    {
        /// <summary>
        /// Получить <see cref="SharedStringItem"/> по его ID
        /// </summary>
        /// <param name="wbPart">Элемент <see cref="WorkbookPart"/></param>
        /// <param name="itemId">ID элемента <see cref="SharedStringItem"/></param>
        /// <returns>Элемент <see cref="SharedStringItem"/> с указанным ID</returns>
        internal static SharedStringItem GetSharedStringItem(this WorkbookPart wbPart, int itemId)
        {
            if (wbPart == null) { throw new ArgumentNullException("wbPart is null"); }
            return wbPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(itemId);
        }


    }
}
