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
        /// Получить таблицу стилей ячеек
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <returns>Объект таблицы стилей ячеек</returns>
        internal static CellFormats GetCellFormatsOf(Stylesheet stylesheet)
        {
            if (stylesheet.CellFormats == null)
            {
                stylesheet.CellFormats = new CellFormats(new CellFormat()) { Count = 1 }; // if not exists, then create blank cell format list
                stylesheet.CellFormats.AppendChild(new CellFormat()); // empty one for index 0, seems to be required
            }
            return stylesheet.CellFormats;
        }
        /// <summary>
        /// Вставить стиль ячейки используя класс CellFormat
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <param name="format">Объект формата ячейки, содержащии информицию о стиле ячейки.</param>
        /// <returns>ID вставленнго формата ячейки в структуре документа.</returns>
        public static uint CellFormat(this Stylesheet stylesheet, CellFormat format)
        {
            var cellFormats = GetCellFormatsOf(stylesheet);
            var formatIndex = cellFormats.MakeSame(format);
            cellFormats.Count = (uint)cellFormats.Elements().Count();
            return formatIndex;
        }

        /// <summary>
        /// Получть стиль ячейки по его ID
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <param name="formatIndex">ID формата ячейки</param>
        /// <returns>Возвращает объект стиля ячейки</returns>
        public static CellFormat CellFormat(this Stylesheet stylesheet, uint formatIndex)
        {
            return GetCellFormatsOf(stylesheet).Elements<CellFormat>().ElementAt((int)formatIndex);
        }

        /// <summary>
        /// Получть стиль ячейки по его ID
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <param name="formatIndex">ID формата ячейки</param>
        /// <returns>Возвращает объект стиля ячейки</returns>
        public static CellFormat CellFormat(this Stylesheet stylesheet, int formatIndex)
        {
            return GetCellFormatsOf(stylesheet).Elements<CellFormat>().ElementAt(formatIndex);
        }

    }
}
