using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class StylesheetCellFormat
    {
        /// <summary>
        /// Вставить стиль ячейки используя класс CellFormat
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <param name="format">Объект формата ячейки, содержащии информицию о стиле ячейки.</param>
        /// <returns>ID вставленнго формата ячейки в структуре документа.</returns>
        public static uint CellFormat(this Stylesheet stylesheet, CellFormat format)
        {
            return stylesheet.GetCellFormats().CellFormat(format);
        }

        /// <summary>
        /// Получть стиль ячейки по его ID
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <param name="formatIndex">ID формата ячейки</param>
        /// <returns>Возвращает объект стиля ячейки</returns>
        public static CellFormat CellFormat(this Stylesheet stylesheet, uint formatIndex)
        {
            return stylesheet.GetCellFormats().CellFormat(formatIndex);
        }

        /// <summary>
        /// Получть стиль ячейки по его ID
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <param name="formatIndex">ID формата ячейки</param>
        /// <returns>Возвращает объект стиля ячейки</returns>
        public static CellFormat CellFormat(this Stylesheet stylesheet, int formatIndex)
        {
            return stylesheet.GetCellFormats().CellFormat(formatIndex);
        }

    }
}
