using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class CellFormatsCellFormat
    {
        /// <summary>
        /// Получть стиль ячейки по его ID
        /// </summary>
        /// <param name="cellFormats">Таблица форматов ячеек</param>
        /// <param name="formatIndex">ID формата ячейки</param>
        /// <returns>Возвращает объект стиля ячейки</returns>
        public static CellFormat CellFormat(this CellFormats cellFormats, int formatIndex)
        {
            return cellFormats.Elements<CellFormat>().ElementAt(formatIndex);
        }

        /// <summary>
        /// Получть стиль ячейки по его ID
        /// </summary>
        /// <param name="cellFormats">Таблица форматов ячеек</param>
        /// <param name="formatIndex">ID формата ячейки</param>
        /// <returns>Возвращает объект стиля ячейки</returns>
        public static CellFormat CellFormat(this CellFormats cellFormats, uint formatIndex)
        {
            return cellFormats.CellFormat((int)formatIndex);
        }

        /// <summary>
        /// Вставить стиль ячейки используя класс CellFormat
        /// </summary>
        /// <param name="cellFormats">Таблица форматов ячеек</param>
        /// <param name="format">Объект формата ячейки, содержащии информицию о стиле ячейки.</param>
        /// <returns>ID вставленнго формата ячейки в структуре документа.</returns>
        public static uint CellFormat(this CellFormats cellFormats, CellFormat format)
        {
            var formatIndex = cellFormats.MakeSame(format);
            cellFormats.Count = (uint)cellFormats.Elements().Count();
            return formatIndex;
        }
    }
}