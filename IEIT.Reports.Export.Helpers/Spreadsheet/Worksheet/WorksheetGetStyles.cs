using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetGetStyles
    {
        /// <summary>
        /// Получение стилей по перечислениям
        /// </summary>
        /// <typeparam name="T">Перечисление содержащее названия стилей</typeparam>
        /// <param name="worksheet">Лист в котором содержатся все стили с названиями соответствующие перечислению</param>
        /// <returns>Именнованный массив со значением указанного перечисления в виде ключа, и с индексом стиля в виде значения</returns>
        public static IDictionary<T, UInt32Value> GetStylesOf<T>(this Worksheet worksheet)
        {
            var dict = new Dictionary<T, UInt32Value>();
            foreach (var pair in worksheet.GetStylesOf(typeof(T)))
            {
                dict.Add((T)pair.Key, pair.Value);
            }
            return dict;
        }

        /// <summary>
        /// Формат регулярного выражения для поиска по названиям стилей
        /// </summary>
        internal const string RGX_NamedStylecellValueFormat = "^({0}\\.)?{1}$";

        /// <summary>
        /// Получение стилей по перечислениям
        /// </summary>
        /// <param name="enum">Перечисление содержащее названия стилей</param>
        /// <param name="worksheet">Лист в котором содержатся все стили с названиями соответствующие перечислению</param>
        /// <returns>Именнованный массив со значением указанного перечисления в виде ключа, и с индексом стиля в виде значения</returns>
        public static IDictionary<object, UInt32Value> GetStylesOf(this Worksheet worksheet, Type @enum)
        {
            if (!@enum.IsEnum)
            {
                throw new Exception($"Передаваемый тип в метод (расширение) Worksheet.GetStylesOf() должен быть Enum");
            }

            var dict = new Dictionary<object, UInt32Value>();
            var typeName = @enum.Name;
            //Пробегаемся по всем элементам в списке
            foreach (var val in Enum.GetValues(@enum))
            {
                var styleName = Enum.GetName(@enum, val);
                var rgxPattern = string.Format(RGX_NamedStylecellValueFormat, Regex.Escape(typeName), Regex.Escape(styleName));
                var rgx = new Regex(rgxPattern);
                //Находим нужную ячейку со стилем
                var cell = worksheet.FindCells(rgx).FirstOrDefault();
                //Записываем индекс стиля в массив
                dict.Add(val, cell?.StyleIndex);
            }
            return dict;
        }

        /// <summary>
        /// Получить стили по адресам ячеек
        /// </summary>
        /// <param name="worksheet">Лист в котором содержатся необходимые стили</param>
        /// <returns>Именнованный массив с адресом ячейки в виде ключа, и с индексом стиля этой ячейки в виде значения</returns>
        public static IDictionary<string, UInt32Value> GetStyles(this Worksheet worksheet)
        {
            var dict = new Dictionary<string, UInt32Value>();
            foreach (var cell in worksheet.Descendants<Cell>())
            {
                if (cell.CellReference == null && cell.CellReference.HasValue) { continue; } //TODO: cell.Address()
                dict.Add(cell.CellReference, cell.StyleIndex);
            }
            return dict;
        }
    }
}
