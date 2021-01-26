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
        /// Получить таблицу форматировании текста. Создает, если такого нет.
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <returns>возвращает таблицу форматировании текста</returns>

        internal static Fonts GetFontsOf(Stylesheet stylesheet)
        {
            if (stylesheet.Fonts == null) { stylesheet.Fonts = new Fonts(new Font()) { Count = 1 }; } // blank font list, if not exists
            return stylesheet.Fonts;
        }
        /// <summary>
        /// Получить формат текста
        /// </summary>
        /// <param name="stylesheet">таблица стилей</param>
        /// <param name="fontIndex">индекс формата текста</param>
        /// <returns>возвращает объект формата текста</returns>
        public static Font Font(this Stylesheet stylesheet, int fontIndex)
        {
            return GetFontsOf(stylesheet).Elements<Font>().ElementAt(fontIndex);
        }
        /// <summary>
        /// Получить формат текста
        /// </summary>
        /// <param name="stylesheet">таблица стилей</param>
        /// <param name="fontIndex">индекс формата текста</param>
        /// <returns>возвращает объект формата текста</returns>
        public static Font Font(this Stylesheet stylesheet, uint fontIndex)
        {
            return GetFontsOf(stylesheet).Elements<Font>().ElementAt((int)fontIndex);
        }

        /// <summary>
        /// Получить формат текста
        /// <para>Depricated: use <see cref="StylesheetFont.Font(Stylesheet, uint)"/> instead</para>
        /// </summary>
        /// <param name="stylesheet">таблица стилей</param>
        /// <param name="fontIndex">индекс формата текста</param>
        /// <returns>возвращает объект формата текста</returns>
        public static Font GetFont(this Stylesheet stylesheet, uint fontIndex)
        {
            return GetFontsOf(stylesheet).Elements<Font>().ElementAt((int)fontIndex);
        }

        /// <summary>
        /// Вставить формат текста
        /// </summary>
        /// <param name="stylesheet">таблица стилей</param>
        /// <param name="font">объект формата текста</param>
        /// <returns>возвращает индекс формата текста</returns>
        public static uint Font(this Stylesheet stylesheet, Font font)
        {
            var fontsList = GetFontsOf(stylesheet);
            var fontIndex = fontsList.MakeSame(font);
            fontsList.Count = (uint)fontsList.Elements().Count();
            return fontIndex;
        }

    }
}
