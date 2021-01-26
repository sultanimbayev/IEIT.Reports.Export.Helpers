﻿using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    //TODO: Remove this class
    public static class BordersGetBorder
    {
        /// <summary>
        /// Получить объект стиля границ ячейки
        /// <para>Depricated: use <see cref="StylesheetBorder.Border(Stylesheet, int)"/> instead</para>
        /// </summary>
        /// <param name="borders">Оъект содержащий элементы стлия границ ячеек</param>
        /// <param name="borderIndex">Индекс объекта</param>
        /// <returns>Объект границ ячейки</returns>
        public static Border GetBorder(this Borders borders, int borderIndex)
        {
            var border = borders.Elements().ElementAt(borderIndex) as Border;
            return border;
        }

        /// <summary>
        /// Получить объект стиля границ ячейки
        /// <para>Depricated: use <see cref="StylesheetBorder.Border(Stylesheet, uint)"/> instead</para>
        /// </summary>
        /// <param name="borders">Оъект содержащий элементы стиля границ ячеек</param>
        /// <param name="borderIndex">Индекс объекта</param>
        /// <returns>Объект границ ячейки</returns>
        public static Border GetBorder(this Borders borders, uint borderIndex)
        {
            return borders.GetBorder((int)borderIndex);
        }

    }
}
