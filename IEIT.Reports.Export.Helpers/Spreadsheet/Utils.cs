using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Configuration;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class Utils
    {
        /// <summary>
        /// Название параметра в App.Config или Web.Config в котором хранится
        /// путь к директории с XML элементами необходимые для работы всех расширении
        /// </summary>
        private const string CONFIG_KEY_ELEMENTS_PATH = "OpenXMLElementsPath";

        /// <summary>
        /// Путь к директории с XML элементами
        /// </summary>
        private static string ElementsDir { get { return ConfigurationManager.AppSettings[CONFIG_KEY_ELEMENTS_PATH]; } }
        
        /// <summary>
        /// Полный путь к папке, где лежит DLL с данной библиотекой
        /// </summary>
        private static string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return System.IO.Path.GetDirectoryName(path);
            }
        }


        /// <summary>
        /// Получить полный путь относительно папки где лежит DLL с данной библиотекой
        /// </summary>
        /// <param name="relativePath"></param>
        /// <returns></returns>
        private static string GetFullPath(string relativePath)
        {
            return System.IO.Path.GetFullPath($"{AssemblyDirectory}\\{relativePath}");
        }

        /// <summary>
        /// Получить номер строки из адреса ячейки
        /// </summary>
        /// <param name="address">Адрес ячейки</param>
        /// <returns></returns>
        public static uint ToRowNum(string address)
        {
            uint result = 0;

            var value = address.TrimStart("ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray());
            
            if (uint.TryParse(value, out result)){}
            return result;
        }
        
        /// <summary>
        /// Получить номер колонки (начиная с 1-го)
        /// </summary>
        /// <param name="address">Адрес ячейки или индекс колонки</param>
        /// <returns>Номер колонки</returns>
        public static uint ToColumNum(string address)
        {
            var value = address.TrimEnd("1234567890".ToCharArray());
            var digits = value.PadLeft(3).Select(x => "ABCDEFGHIJKLMNOPQRSTUVWXYZ".IndexOf(x));
            return (uint)digits.Aggregate(0, (current, index) => (current * 26) + (index + 1));
        }

        /// <summary>
        /// Получить индекс колонки в буквенном значении
        /// </summary>
        /// <param name="address">Номер колонки (начиная с 1-го) или номер колонки</param>
        /// <returns>Индекс колонки в буквенном значении</returns>
        /// <example>Utils.ToColumnName("C14") => "С" или Utils.ToColumnName("5") => "E"</example>
        public static string ToColumnName(string address)
        {
            if (IsNumber(address)) { return ToColumnName(uint.Parse(address)); }
            return address.TrimEnd("1234567890".ToCharArray());
        }

        /// <summary>
        /// Получить индекс колонки в буквенном значении
        /// </summary>
        /// <param name="columnNumber">Номер колонки (начиная с 1-го). Значение должно быть больше нуля.</param>
        /// <returns>Индекс колонки в виде буквы латинского языка</returns>
        /// <example>Utils.ToColumnName(2) => "B"</example>
        public static string ToColumnName(int columnNumber)
        {
            return ToColumnName((uint)columnNumber);
        }

        /// <summary>
        /// Получить индекс колонки в буквенном значении
        /// </summary>
        /// <param name="columnNumber">Номер колонки (начиная с 1-го). Значение должно быть больше нуля.</param>
        /// <returns>Индекс колонки в виде буквы латинского языка</returns>
        /// <example>Utils.ToColumnName(2) => "B"</example>
        public static string ToColumnName(uint columnNumber)
        {
            int dividend = (int)columnNumber;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }


        /// <summary>
        /// Узнать, является ли данная строка адресом ячейки
        /// </summary>
        /// <param name="value">Проверяемая строка</param>
        /// <returns>true если данная строка является валидным адресом ячейки, false в обратном случае</returns>
        public static bool IsCellAddress(this string value)
        {
            if (value == null) { return false; }
            value = value.TrimStart("ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray());
            if (value.StartsWith("0")) { return false; }
            value = value.TrimEnd("1234567890".ToCharArray());
            return value.Equals(string.Empty);
        }


        /// <summary>
        /// Получить XML по ID элемента
        /// </summary>
        /// <param name="id">ID элемента, путь к XML файлу относительно папки с XML элеметнами</param>
        /// <returns>Контент XML файла по указанному пути</returns>
        private static string GetDefaultElement(string id)
        {
            var dir = GetFullPath(ElementsDir);

            if (!id.EndsWith(".xml"))
            {
                id = id + ".xml";
            }

            var filepath = $"{dir}\\{id}";
            string xmlString = System.IO.File.ReadAllText(filepath);
            return xmlString;
        }

        /// <summary>
        /// Создает елемент из указанного файла
        /// </summary>
        /// <typeparam name="T">Тип элемента</typeparam>
        /// <param name="element">Элемент, который будет переопределен элементом из файла</param>
        /// <param name="id">ID элемента, путь к XML файлу относительно папки с XML элеметнами</param>
        /// <returns></returns>
        public static T From<T>(this T element, string id) where T: OpenXmlElement
        {
            var newElemStr = GetDefaultElement(id);
            var elementType = typeof(T);
            return element = Activator.CreateInstance(elementType, new object[] { newElemStr }) as T;
        }


        /// <summary>
        /// Получить первого потомка с указанным типом
        /// </summary>
        /// <typeparam name="T">Тип искомого потомка</typeparam>
        /// <param name="element">Родительский элемент, потомок которого нужно найти</param>
        /// <returns>
        /// Потомок типа <typeparamref name="T"/> для элемента <paramref name="element"/>
        /// или null если потомок не найден
        /// </returns>
        public static T FirstDescendant<T>(this OpenXmlElement element) where T : OpenXmlElement
        {
            return element.Descendants<T>().FirstOrDefault();
        }

        /// <summary>
        /// Заменяет один элемент в древе другим элементом.
        /// Не требует отвязки элементов от их древа, и не удаляет
        /// заменяющий элемент из его дерева.
        /// </summary>
        /// <typeparam name="T1">Тип заменяемого элемента</typeparam>
        /// <typeparam name="T2">Тип заменяющего элемента</typeparam>
        /// <param name="oldElement">Заменяемый элемент</param>
        /// <param name="newElement">Заменяющий элемент</param>
        /// <returns>Новый элемент после замены</returns>
        public static T2 ReplaceBy<T1, T2>(this T1 oldElement, T2 newElement) where T1 : OpenXmlElement where T2 : OpenXmlElement
        {
            if(oldElement == null) { throw new ArgumentNullException("oldElement"); }
            if(newElement == null) { throw new ArgumentNullException("newElement"); }

            var parent = oldElement.Parent;
            if (parent != null) {
                var replacedElement = parent.ReplaceChild((newElement = newElement.CloneNode(true) as T2), oldElement) as T1;
                return replacedElement.Equals(oldElement) ? newElement : null;
            }
            return null;
        }
        

        /// <summary>
        /// Проверяет, является ли значение числом
        /// </summary>
        /// <param name="value">Проверяемое значение</param>
        /// <returns>true если значение числовое, false в обратном случае</returns>
        public static bool IsNumber(object value)
        {
            double d;
            long l;
            decimal dec;
            var valStr = value.ToString() as string;
            return double.TryParse(valStr, out d) || long.TryParse(valStr, out l) || decimal.TryParse(valStr, out dec);
        }

        /// <summary>
        /// Преобразует значение в объект <see cref="Drawing.Text"/>
        /// </summary>
        /// <param name="str">Преобразуемое значение</param>
        /// <returns>Объект типа <see cref="Drawing.Text"/> с исходным значением <paramref name="str"/></returns>
        public static Drawing.Text ToDrawingText(this string str)
        {
            return new Drawing.Text(str);
        }

        /// <summary>
        /// Преобразует значение в объект <see cref="Text"/>
        /// </summary>
        /// <param name="str">Преобразуемое значение</param>
        /// <returns>Объект типа <see cref="Text"/> с исходным значением <paramref name="str"/></returns>
        public static Text ToText(this string str)
        {
            return new Text(str);
        }

        /// <summary>
        /// Получить адреса ячеек из адреса ряда ячеек
        /// </summary>
        /// <param name="cellsRange">Адрес ряда ячеек</param>
        /// <returns>Массив из адресов ячеек находящиеся в указанном промежутке</returns>
        public static List<string> CellAddressesFrom(string cellsRange)
        {
            Regex rgxSingle = new Regex(Common.RGX_PAT_CA);
            Regex rgxRange = new Regex(Common.RGX_PAT_CA_RANGE);
            if (rgxSingle.IsMatch(cellsRange)) { return new List<string>(){ cellsRange }; }
            if (!rgxRange.IsMatch(cellsRange)) { throw new FormatException($"Не удалось считать адреса ячеек {cellsRange}"); }

            var addrs = cellsRange.Split(':');
            uint initCol = ToColumNum(addrs[0].ToUpper());
            uint initRow = ToRowNum(addrs[0].ToUpper());

            uint finalCol = ToColumNum(addrs[1].ToUpper());
            uint finalRow = ToRowNum(addrs[1].ToUpper());

            List<string> cellAddrs = new List<string>();

            for (uint row = initRow; row <= finalRow; row++)
            {
                for (uint col = initCol; col <= finalCol; col++)
                {
                    cellAddrs.Add(ToColumnName(col) + row.ToString());
                }
            }

            return cellAddrs;
        }
        
        /// <summary>
        /// Заменяет все вхождения одной строки другой строкой
        /// </summary>
        /// <typeparam name="T">Тип элемента <see cref="OpenXmlLeafTextElement"/></typeparam>
        /// <param name="formula">Элемент, текст которого будет преобразован</param>
        /// <param name="oldValue">Заменяемое значение</param>
        /// <param name="newValue">Заменяющее значение</param>
        /// <returns>Исходный элемент с измененным значением <paramref name="formula"/></returns>
        public static T Replace<T>(this T formula, string oldValue, string newValue) where T : OpenXmlLeafTextElement
        {
            if(formula.Text == null) { return formula; }
            formula.Text = formula.Text.Replace(oldValue, newValue);
            return formula;
        }

        /// <summary>
        /// Заменяет все вхождения регулярного выражения
        /// </summary>
        /// <typeparam name="T">Тип элемента <see cref="OpenXmlLeafTextElement"/></typeparam>
        /// <param name="formula">Элемент, текст которого будет преобразован</param>
        /// <param name="pattern">Искомое регулярное выражение</param>
        /// <param name="replacement">Заменяющее значение</param>
        /// <param name="options">Дополнительные параметры регулярного выражения</param>
        /// <returns>Исходный элемент с измененным значением <paramref name="formula"/></returns>
        public static T RegexReplace<T>(this T formula, string pattern, string replacement, RegexOptions options = RegexOptions.None) where T : OpenXmlLeafTextElement
        {
            Regex regEx = new Regex(pattern, options);
            formula.Text = regEx.Replace(formula.Text, replacement);
            return formula;
        }

        /// <summary>
        /// Заменяет все вхождения регулярного выражения в дочерних элементах типа <typeparamref name="T"/>
        /// </summary>
        /// <typeparam name="T">Тип элемента <see cref="OpenXmlLeafTextElement"/></typeparam>
        /// <param name="element">Родительский элемент, дочерние объекты которого будут преобразованы</param>
        /// <param name="pattern">Искомое регулярное выражение</param>
        /// <param name="replacement">Заменяющее значение</param>
        /// <param name="options">Дополнительные параметры регулярного выражения</param>
        public static void RegexReplaceIn<T>(this OpenXmlElement element, string pattern, string replacement, RegexOptions options = RegexOptions.None) where T : OpenXmlLeafTextElement
        {
            Regex regEx = new Regex(pattern, options);
            var formulas = element.Descendants<T>();
            foreach (var formula in formulas)
            {
                formula.Text = regEx.Replace(formula.Text, replacement);
            }
        }

        /// <summary>
        /// Заменяет все вхождения одной строки другой строкой в дочерних 
        /// элементах типа <typeparamref name="T"/> данного объекта 
        /// </summary>
        /// <typeparam name="T">Тип элемента <see cref="OpenXmlLeafTextElement"/></typeparam>
        /// <param name="element">Родительский элемент, дочерние объекты которого будут преобразованы</param>
        /// <param name="oldValue">Заменяемое значение</param>
        /// <param name="newValue">Заменяющее значение</param>
        public static void ReplaceIn<T>(this OpenXmlElement element, string oldValue, string newValue) where T : OpenXmlLeafTextElement
        {
            var formulas = element.Descendants<T>();
            foreach(var formula in formulas)
            {
                formula.Replace(oldValue, newValue);
            }
        }

        /// <summary>
        /// Определяет, являются ли два элемента схожими.
        /// Даже если два элемента null он определяет их как одинаковые.
        /// Не зависит от положения в древе так как используется 
        /// метод <see cref="OpenXmlElement.CloneNode(bool)"/> перед сравнением
        /// </summary>
        /// <typeparam name="T">Тип сравниваемых элементов</typeparam>
        /// <param name="source">Первое сравниваемый элемент</param>
        /// <param name="target">Второй сравниваемый элемент</param>
        /// <returns>true если два элемента схожи, false в обратном случае</returns>
        public static bool SameAs<T>(this T source, T target) where T : OpenXmlElement
        {
            if(source == null && target == null) { return true; }
            if(source == null || target == null) { return false; }
            return source.CloneNode(true).Equals(target.CloneNode(true));
        }
        

        /// <summary>
        /// Получить ближайшего родителя типа <typeparamref name="T"/>
        /// </summary>
        /// <typeparam name="T">Тип искомого родителя</typeparam>
        /// <param name="element">Элемент, родителя которого требуется найти</param>
        /// <returns>Ближайший родительский элемент типа <typeparamref name="T"/></returns>
        public static T ParentOfType<T>(this OpenXmlElement element) where T : OpenXmlElement
        {
            OpenXmlElement parent;
            while((parent = element.Parent) != null)
            {
                if (element.Parent == null) { return null; }
                if (element.Parent is T) { return element.Parent as T; }
                element = parent;
            }
            return null;
        }


        /// <summary>
        /// Получить ближайшего родительской части документа типа <typeparamref name="T"/>
        /// </summary>
        /// <typeparam name="T">Тип искомого родителя</typeparam>
        /// <param name="part">Часть документа, предка котого требуется найти</param>
        /// <returns>Ближайший предок типа <typeparamref name="T"/></returns>
        public static T ParentPartOfType<T>(this OpenXmlPart part) where T : OpenXmlPart
        {
            var parentParts = part.GetParentParts();
            if(parentParts == null || parentParts.Count() == 0) { return null; }
            foreach(var p in parentParts)
            {
                if(p == null) { continue; }
                if(!(p is T))
                {
                    var p2 = p.ParentPartOfType<T>();
                    if(p2 == null) { continue; }
                    return p2;
                }
                return p as T;
            }
            return null;
        }

        /// <summary>
        /// Получить индекс элемента в родительском списке, начинается с нуля.
        /// </summary>
        /// <param name="element">Элемент индекс которого нужно найти</param>
        /// <returns>Индекс данного элемента среди элементов родителя, начинается с нуля.</returns>
        public static int Index(this OpenXmlElement element) 
        {
            if(element == null) { throw new ArgumentNullException("Cannot get index of null element!"); }
            if(element.Parent == null) { throw new MissingMemberException("Cannot get index of element that doesn't have parent!"); }
            return element.Parent.Elements().ToList().FindIndex(el => el.Equals(element));
        }
        
    }
}
