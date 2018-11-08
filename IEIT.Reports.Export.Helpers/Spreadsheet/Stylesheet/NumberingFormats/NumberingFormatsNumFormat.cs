using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class NumberingFormatsNumFormat
    {
        
        /// <summary>
        /// Получить формат числа
        /// </summary>
        /// <param name="numFormats">таблица форматов чисел</param>
        /// <param name="formatIndex">индекс формата числа</param>
        /// <returns>возвращает объект формата числа</returns>
        public static NumberingFormat NumFormat(this NumberingFormats numFormats, int formatIndex)
        {
            return numFormats.Elements<NumberingFormat>().ElementAt(formatIndex);
        }

        /// <summary>
        /// Получить формат числа
        /// </summary>
        /// <param name="numFormats">таблица форматов чисел</param>
        /// <param name="formatIndex">индекс формата числа</param>
        /// <returns>возвращает объект формата числа</returns>
        public static NumberingFormat NumFormat(this NumberingFormats numFormats, uint formatIndex)
        {
            return numFormats.NumFormat((int)formatIndex);
        }

        /// <summary>
        /// Вставить формат числа в структуру документа
        /// </summary>
        /// <param name="numFormats">таблица форматов чисел</param>
        /// <param name="numFormat">формат числа</param>
        /// <returns>Индекс вставленного формата числа</returns>
        public static uint NumFormat(this NumberingFormats numFormats, NumberingFormat numFormat)
        {
            var formatIndex = numFormats.MakeSame(numFormat);
            numFormats.Count = (uint)numFormats.Elements().Count();
            return formatIndex;
        }
    }
}
