using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{

    /// <summary>
    /// Удаление элементов
    /// </summary>
    public static class _RemoveElements
    {
        /// <summary>
        /// Удаляет элементы
        /// </summary>
        /// <param name="elements">Элементы которые будут удалены</param>
        /// <param name="deleteSectionIfEmpty">
        ///     Если указан как true, то удаляется также родительские 
        ///     элементы когда они остаются пустыми.
        /// </param>
        public static void RemoveElements(IEnumerable<OpenXmlElement> elements, bool deleteSectionIfEmpty = false)
        {
            OpenXmlElement prevItem = null;
            int idx = 0;
            while (elements.Count() != idx)
            {
                var item = elements.ElementAtOrDefault(idx);
                if (item == null || item.Equals(prevItem))
                {
                    idx++;
                    continue;
                }
                if (item == default(OpenXmlElement))
                {
                    break;
                }
                var par = item.Parent;
                item.Remove();
                if (deleteSectionIfEmpty && par != null && par.ChildElements.Count == 0)
                {
                    par.Remove();
                }
                prevItem = item;
            }
        }
    }
}
