using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class StylesheetGetNumFormats
    {
        /// <summary>
        /// Получить таблицу формата чисел
        /// </summary>
        /// <param name="stylesheet">таблица стилей</param>
        /// <returns>возвращает таблицу формата чисел</returns>
        internal static NumberingFormats GetNumFormats(this Stylesheet stylesheet)
        {
            if (stylesheet.NumberingFormats == null) { stylesheet.NumberingFormats = new NumberingFormats() { Count = 0 }; }
            return stylesheet.NumberingFormats;
        }
    }
}
