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
        public const int BUILTIN_NUMFORMATS_COUNT = 164;
        /// <summary>
        /// Получить таблицу формата чисел
        /// </summary>
        /// <param name="stylesheet">таблица стилей</param>
        /// <returns>возвращает таблицу формата чисел</returns>
        internal static NumberingFormats GetNumFormatsOf(Stylesheet stylesheet)
        {
            if (stylesheet.NumberingFormats == null) { stylesheet.NumberingFormats = new NumberingFormats() { Count = 0 }; }
            return stylesheet.NumberingFormats;
        }
        /// <summary>
        /// Получить формат числа
        /// </summary>
        /// <param name="stylesheet">таблица стилей</param>
        /// <param name="numFormatId">индекс формата числа</param>
        /// <returns>возвращает объект формата числа</returns>
        public static NumberingFormat NumFormat(this Stylesheet stylesheet, int numFormatId)
        {
            if(numFormatId < BUILTIN_NUMFORMATS_COUNT) { return null; }
            return GetNumFormatsOf(stylesheet).Elements<NumberingFormat>().ElementAt(numFormatId - BUILTIN_NUMFORMATS_COUNT);
        }


        /// <summary>
        /// Получить формат числа
        /// </summary>
        /// <param name="stylesheet">таблица стилей</param>
        /// <param name="numFormatId">индекс формата числа</param>
        /// <returns>возвращает объект формата числа</returns>
        public static NumberingFormat NumFormat(this Stylesheet stylesheet, uint numFormatId)
        {
            if (numFormatId < BUILTIN_NUMFORMATS_COUNT) { return null; }
            return GetNumFormatsOf(stylesheet).Elements<NumberingFormat>().ElementAt((int)numFormatId - BUILTIN_NUMFORMATS_COUNT);
        }


        /// <summary>
        /// Вставить формат числа в структуру документа
        /// </summary>
        /// <param name="stylesheet">таблица стилей</param>
        /// <param name="numFormat">формат числа</param>
        /// <returns>Индекс вставленного формата числа</returns>
        public static uint NumFormat(this Stylesheet stylesheet, NumberingFormat numFormat)
        {
            var numFormats = GetNumFormatsOf(stylesheet);
            var numFormatId = numFormats.MakeSame(numFormat) + BUILTIN_NUMFORMATS_COUNT;
            numFormats.Count = (uint)numFormats.Elements().Count();
            return numFormatId;
        }
    }
}
