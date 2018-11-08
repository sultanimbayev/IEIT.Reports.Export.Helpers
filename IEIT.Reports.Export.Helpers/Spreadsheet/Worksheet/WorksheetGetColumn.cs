using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetGetColumn
    {
        /// <summary>
        /// Получить объект для работы со столбцом, с указанным адресом.
        /// </summary>
        /// <param name="worksheet">Объект рабочего листа</param>
        /// <param name="columnNumber">Номер запрашиваемого столбца (начиная с 1-го)</param>
        /// <returns>Объект для работы со столбцом</returns>
        public static Models.Column GetColumn(this Worksheet worksheet, int columnNumber)
        {
            return new Models.Column(worksheet, columnNumber);
        }

        /// <summary>
        /// Получить объект для работы со столбцом, с указанным адресом.
        /// </summary>
        /// <param name="worksheet">Объект рабочего листа</param>
        /// <param name="columnNumber">Номер запрашиваемого столбца (начиная с 1-го)</param>
        /// <returns>Объект для работы со столбцом</returns>
        public static Models.Column GetColumn(this Worksheet worksheet, uint columnNumber)
        {
            return new Models.Column(worksheet, columnNumber);
        }

        /// <summary>
        /// Получить объект для работы со столбцом, с указанным адресом.
        /// </summary>
        /// <param name="worksheet">Объект рабочего листа</param>
        /// <param name="columnName">Название запрашиваемого столбца, латинские буквы.</param>
        /// <returns>Объект для работы со столбцом</returns>
        public static Models.Column GetColumn(this Worksheet worksheet, string columnName)
        {
            return new Models.Column(worksheet, columnName);
        }


    }
}
