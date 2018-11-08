using DocumentFormat.OpenXml.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class StylesheetGetBorders
    {
        /// <summary>
        /// Получить объект стиля содержащий элементы границ ячеек
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <returns>Объект содержащий элементы границ ячеек</returns>
        internal static Borders GetBorders(this Stylesheet stylesheet)
        {
            if (stylesheet.Borders == null) { stylesheet.Borders = new Borders(new Border()) { Count = 1 }; } // blank border list, if not exists
            return stylesheet.Borders;
        }

        /// <summary>
        /// Получить объект стиля границ ячейки
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <param name="borderIndex">Индекс объекта границ ячеек</param>
        /// <returns>Объект границ ячейки</returns>
        public static Border GetBorder(this Stylesheet stylesheet, int borderIndex)
        {
            return stylesheet.GetBorders().GetBorder(borderIndex);
        }


        /// <summary>
        /// Получить объект стиля границ ячейки
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <param name="borderIndex">Индекс объекта стиля границ ячеек</param>
        /// <returns>Объект границ ячейки</returns>
        public static Border GetBorder(this Stylesheet stylesheet, uint borderIndex)
        {
            return stylesheet.GetBorders().GetBorder(borderIndex);
        }
    }
}
