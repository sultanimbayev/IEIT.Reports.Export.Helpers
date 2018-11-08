using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetGetNamedStyles
    {

        /// <summary>
        /// Получить стили по содержимым ячеек
        /// </summary>
        /// <param name="worksheet">Лист в котором содержатся необходимые стили</param>
        /// <returns>Именнованный массив с содержимым ячейки в виде ключа, и с индексом стиля этой ячейки в виде значения</returns>
        public static IDictionary<string, UInt32Value> GetNamedStyles(this Worksheet worksheet)
        {
            var dict = new Dictionary<string, UInt32Value>();
            foreach (var cell in worksheet.Descendants<Cell>())
            {
                var cellValue = cell.GetValue();

                //Пропускаем пустые ячейки
                if (string.IsNullOrWhiteSpace(cellValue))
                {
                    continue;
                }

                if (dict.ContainsKey(cellValue))
                {
                    throw new Exception("Значения ячеек в листе должны быть уникальными");
                }

                dict.Add(cell.GetValue(), cell.StyleIndex);
            }
            return dict;
        }
    }
}
