using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public enum MatchOption
    {
        Contains,
        Equals,
        StartsWith,
        EndsWith
    }

    public static class WorksheetFindCells
    {


        /// <summary>
        /// Найти ячейки по его содержанию
        /// </summary>
        /// <param name="worksheet">Рабочий лист документа в котором ведется поиск</param>
        /// <param name="searchRgx">Объект регулярного выражения для поиска</param>
        /// <returns>Ячейки содержание которых совпадает с данным выражением</returns>
        public static IEnumerable<Cell> FindCells(this Worksheet worksheet, Regex searchRgx)
        {
            return worksheet.Descendants<Cell>().Where(c => { var val = c.GetValue(); return val != null && searchRgx.IsMatch(val); });
        }

        /// <summary>
        /// Найти ячейки по его содержанию
        /// </summary>
        /// <param name="worksheet">Рабочий лист документа в котором ведется поиск</param>
        /// <param name="searchText">Значение которое должно содержать ячейка</param>
        /// <returns>Ячейки содержание которых совпадает с указанным значением</returns>
        public static IEnumerable<Cell> FindCells(this Worksheet worksheet, string searchText, MatchOption match = MatchOption.Contains)
        {
            Func<string, string, bool> matchDeleg;
            switch (match)
            {
                default:
                    throw new NotImplementedException();
                case MatchOption.Contains:
                    matchDeleg = (_cellText, _searchTxt) => { return _cellText.Contains(_searchTxt); };
                    break;
                case MatchOption.Equals:
                    matchDeleg = (_cellText, _searchTxt) => { return _cellText.Equals(_searchTxt); };
                    break;
                case MatchOption.StartsWith:
                    matchDeleg = (_cellText, _searchTxt) => { return _cellText.StartsWith(_searchTxt); };
                    break;
                case MatchOption.EndsWith:
                    matchDeleg = (_cellText, _searchTxt) => { return _cellText.EndsWith(_searchTxt); };
                    break;
            }
            return worksheet.Descendants<Cell>().Where(c => { var val = c.GetValue(); return val != null && matchDeleg(val, searchText); });
        }
    }
}
