using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class CellGetRow
    {
        /// <summary>
        /// Получить строку в которой находится данная ячейка
        /// </summary>
        /// <param name="cell">Объект ячейки OpenXML</param>
        /// <returns>Объект строки, в которой находится ячейка</returns>
        public static Row GetRow(this Cell cell)
        {
            return cell.ParentOfType<Row>();
        }
    }
}
