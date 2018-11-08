using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class StylesheetGetCellFormats
    {
        /// <summary>
        /// Получить таблицу стилей ячеек
        /// </summary>
        /// <param name="stylesheet">Таблица стилей</param>
        /// <returns>Объект таблицы стилей ячеек</returns>
        internal static CellFormats GetCellFormats(this Stylesheet stylesheet)
        {
            if (stylesheet.CellFormats == null)
            {
                stylesheet.CellFormats = new CellFormats(new CellFormat()) { Count = 1 }; // if not exists, then create blank cell format list
                stylesheet.CellFormats.AppendChild(new CellFormat()); // empty one for index 0, seems to be required
            }
            return stylesheet.CellFormats;
        }

    }
}
