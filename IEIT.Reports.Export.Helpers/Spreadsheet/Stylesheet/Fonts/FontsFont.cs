using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class FontsFont
    {
        /// <summary>
        /// Получить формат текста
        /// </summary>
        /// <param name="fonts">таблица форматов текста</param>
        /// <param name="fontIndex">индекс формата текста</param>
        /// <returns>возвращает объект формата текста</returns>
        public static Font Font(this Fonts fonts, int fontIndex)
        {
            return fonts.Elements<Font>().ElementAt(fontIndex);
        }

        /// <summary>
        /// Получить формат текста
        /// </summary>
        /// <param name="fonts">таблица форматов текста</param>
        /// <param name="fontIndex">индекс формата текста</param>
        /// <returns>возвращает объект формата текста</returns>
        public static Font Font(this Fonts fonts, uint fontIndex)
        {
            return fonts.Font((int)fontIndex);
        }

        /// <summary>
        /// Вставить формат текста
        /// </summary>
        /// <param name="fonts">таблица форматов текста</param>
        /// <param name="font">объект формата текста</param>
        /// <returns>возвращает индекс формата текста</returns>
        public static uint Font(this Fonts fonts, Font font)
        {
            var fontIndex = fonts.MakeSame(font);
            fonts.Count = (uint)fonts.Elements().Count();
            return fontIndex;
        }


    }
}
