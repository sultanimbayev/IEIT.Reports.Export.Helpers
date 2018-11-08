using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class StylesheetFont
    {
        /// <summary>
        /// Получить формат текста
        /// </summary>
        /// <param name="stylesheet">таблица стилей</param>
        /// <param name="fontIndex">индекс формата текста</param>
        /// <returns>возвращает объект формата текста</returns>
        public static Font Font(this Stylesheet stylesheet, int fontIndex)
        {
            return stylesheet.GetFonts().Font(fontIndex);
        }

        /// <summary>
        /// Получить формат текста
        /// </summary>
        /// <param name="stylesheet">таблица стилей</param>
        /// <param name="fontIndex">индекс формата текста</param>
        /// <returns>возвращает объект формата текста</returns>
        public static Font GetFont(this Stylesheet stylesheet, uint fontIndex)
        {
            return stylesheet.GetFonts().Font(fontIndex);
        }

        /// <summary>
        /// Вставить формат текста
        /// </summary>
        /// <param name="stylesheet">таблица стилей</param>
        /// <param name="font">объект формата текста</param>
        /// <returns>возвращает индекс формата текста</returns>
        public static uint Font(this Stylesheet stylesheet, Font font)
        {
            return stylesheet.GetFonts().Font(font);
        }

    }
}
