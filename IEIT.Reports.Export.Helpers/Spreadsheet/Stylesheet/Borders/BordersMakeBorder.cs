using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class BordersMakeBorder
    {
        /// <summary>
        /// Создать стиль границы ячеек. Возвращает индекс 
        /// созданного стиля.
        /// Не создает обект если такой стиль уже иммется, и
        /// возвращает индекс уже созданного стиля.
        /// </summary>
        /// <param name="borders">Оъект содержащий элементы стиля границ ячеек</param>
        /// <param name="border">Стиль границ ячейки</param>
        /// <returns>
        /// Возвращает индекс созданного стиля границы, или индекс имееющегося стиля границы.
        /// </returns>
        public static uint MakeBorder(this Borders borders, Border border)
        {
            var borderIndex = borders.MakeSame(border);
            borders.Count = (uint)borders.Elements().Count();
            return borderIndex;
        }

        /// <summary>
        /// Создать стиль границы ячеек. Возвращает индекс 
        /// созданного стиля.
        /// Не создает обект если такой стиль уже иммется, и
        /// возвращает индекс уже созданного стиля.
        /// </summary>
        /// <param name="stylesheet">Таблица границ ячеек</param>
        /// <param name="border">Стиль границ ячеек</param>
        /// <returns>Возвращает индекс созданного стиля, или индекс имееющегося стиля.</returns>
        public static uint MakeBorder(this Stylesheet stylesheet, Border border)
        {
            return stylesheet.GetBorders().MakeBorder(border);
        }
    }
}
