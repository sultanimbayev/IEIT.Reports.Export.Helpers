using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class StylesheetNumFormat
    {

        /// <summary>
        /// Получить формат числа
        /// </summary>
        /// <param name="stylesheet">таблица стилей</param>
        /// <param name="formatIndex">индекс формата числа</param>
        /// <returns>возвращает объект формата числа</returns>
        public static NumberingFormat NumFormat(this Stylesheet stylesheet, int formatIndex)
        {
            return stylesheet.GetNumFormats().NumFormat(formatIndex);
        }


        /// <summary>
        /// Получить формат числа
        /// </summary>
        /// <param name="stylesheet">таблица стилей</param>
        /// <param name="formatIndex">индекс формата числа</param>
        /// <returns>возвращает объект формата числа</returns>
        public static NumberingFormat NumFormat(this Stylesheet stylesheet, uint formatIndex)
        {
            return stylesheet.GetNumFormats().NumFormat(formatIndex);
        }


        /// <summary>
        /// Вставить формат числа в структуру документа
        /// </summary>
        /// <param name="stylesheet">таблица стилей</param>
        /// <param name="numFormat">формат числа</param>
        /// <returns>Индекс вставленного формата числа</returns>
        public static uint NumFormat(this Stylesheet stylesheet, NumberingFormat numFormat)
        {
            return stylesheet.GetNumFormats().NumFormat(numFormat);
        }
    }
}
