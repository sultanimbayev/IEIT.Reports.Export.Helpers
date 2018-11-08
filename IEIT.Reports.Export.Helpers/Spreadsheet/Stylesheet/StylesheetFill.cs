using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class StylesheetFill
    {
        /// <summary>
        /// Получить стиль заливку по ID стиля заливки
        /// </summary>
        /// <param name="fills">Таблица стилей заливки</param>
        /// <param name="fillIndex">Индекс стиля заливки</param>
        /// <returns>Объекь стиля заливки</returns>
        public static Fill Fill(this Fills fills, int fillIndex)
        {
            return fills.Elements<Fill>().ElementAt(fillIndex);
        }

        public static Fill Fill(this Fills fills, uint fillIndex)
        {
            return fills.Fill((int)fillIndex);
        }

        public static Fill Fill(this Stylesheet stylesheet, int fillIndex)
        {
            return stylesheet.GetFills().Fill(fillIndex);
        }

        public static Fill Fill(this Stylesheet stylesheet, uint fillIndex)
        {
            return stylesheet.GetFills().Fill(fillIndex);
        }

        public static uint Fill(this Fills fills, Fill fill)
        {
            var fillIndex = fills.MakeSame(fill);
            fills.Count = (uint)fills.Elements().Count();
            return fillIndex;
        }

        public static uint Fill(this Stylesheet stylesheet, Fill fill)
        {
            return stylesheet.GetFills().Fill(fill);
        }
    }
}
