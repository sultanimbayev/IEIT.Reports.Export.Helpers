using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Spreadsheet.Intents;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class StylesheetMakeCellStyle
    {
        /// <summary>
        /// Вставить стиль ячейки используя класс CellFormat
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <param name="cellFormat">Объект формата ячейки, содержащии информицию о стиле ячейки.</param>
        /// <returns>ID вставленнго формата ячейки в структуре документа.</returns>
        public static uint MakeCellStyle(this Stylesheet stylesheet, CellFormat cellFormat)
        {
            return stylesheet.GetCellFormats().CellFormat(cellFormat);
        }

        /// <summary>
        /// Создать новый стиль ячейки
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <returns>"Намерение" для создания стиля ячейки</returns>
        public static MakeStyleIntent MakeCellStyle(this Stylesheet stylesheet)
        {
            return new MakeStyleIntent(stylesheet);
        }

    }
}
