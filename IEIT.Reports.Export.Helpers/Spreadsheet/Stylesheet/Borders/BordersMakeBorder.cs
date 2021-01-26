using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{

    //TODO: Remove this class
    public static class BordersMakeBorder
    {
        /// <summary>
        /// Создать стиль границы ячеек. Возвращает индекс 
        /// созданного стиля.
        /// Не создает обект если такой стиль уже иммется, и
        /// возвращает индекс уже созданного стиля.
        /// <para>Depricated: use <see cref="StylesheetBorder.Border(Stylesheet, Border)"/> instead</para>
        /// </summary>
        /// <param name="stylesheet">Таблица границ ячеек</param>
        /// <param name="border">Стиль границ ячеек</param>
        /// <returns>Возвращает индекс созданного стиля, или индекс имееющегося стиля.</returns>
        public static uint MakeBorder(this Stylesheet stylesheet, Border border)
        {
            return stylesheet.Border(border);
        }
    }
}
