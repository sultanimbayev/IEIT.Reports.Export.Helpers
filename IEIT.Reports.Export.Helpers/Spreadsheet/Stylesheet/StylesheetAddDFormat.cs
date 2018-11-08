using DocumentFormat.OpenXml.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class StylesheetAddDFormat
    {
        /// <summary>
        /// Добавить формат для условного форматирования ячеек.
        /// Возвращает индекс добавленного формата.
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <param name="format">Новый формат</param>
        /// <returns>Индекс добавленного формата</returns>
        public static uint AddDFormat(this Stylesheet stylesheet, DifferentialFormat format)
        {
            if (stylesheet.DifferentialFormats == null)
            {
                stylesheet.DifferentialFormats = new DifferentialFormats() { Count = 0 };
            }
            return stylesheet.DifferentialFormats.AddDFormat(format);
        }

    }
}
