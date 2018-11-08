using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class StylesheetGetFonts
    {
        /// <summary>
        /// Получить таблицу форматировании текста. Создает, если такого нет.
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <returns>возвращает таблицу форматировании текста</returns>
        internal static Fonts GetFonts(this Stylesheet stylesheet)
        {
            if (stylesheet.Fonts == null) { stylesheet.Fonts = new Fonts(new Font()) { Count = 1 }; } // blank font list, if not exists
            return stylesheet.Fonts;
        }
    }
}
