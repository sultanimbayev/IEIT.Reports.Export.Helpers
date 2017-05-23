using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Configuration;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class Utils
    {
        private const string CONFIG_KEY_ELEMENTS_PATH = "OpenXMLElementsPath";
        private static string ElementsDir;
        
        static Utils()
        {
            ElementsDir = ConfigurationManager.AppSettings[CONFIG_KEY_ELEMENTS_PATH];
        }

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

        private static string GetFullPath(string relativePath)
        {
            return System.IO.Path.GetFullPath($"{AssemblyDirectory}\\{relativePath}");
        }

        public static uint ToRowNum(string address)
        {
            uint result = 0;

            var value = address.TrimStart("ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray());
            
            if (uint.TryParse(value, out result)){}
            return result;
        }
        
        
        public static uint ToColumNum(string address)
        {
            var value = address.TrimEnd("1234567890".ToCharArray());
            var digits = value.PadLeft(3).Select(x => "ABCDEFGHIJKLMNOPQRSTUVWXYZ".IndexOf(x));
            return (uint)digits.Aggregate(0, (current, index) => (current * 26) + (index + 1));
        }
        
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

        public static T From<T>(this T element, string id) where T: OpenXmlElement
        {
            var newElemStr = GetDefaultElement(id);
            var elementType = typeof(T);
            return element = Activator.CreateInstance(elementType, new object[] { newElemStr }) as T;
        }

        public static T FirstDescendant<T>(this OpenXmlElement element) where T : OpenXmlElement
        {
            return element.Descendants<T>().FirstOrDefault();
        }

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

        public static bool IsCellAddress(this string value)
        {
            if(value == null) { return false; }
            value = value.TrimStart("ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray());
            if (value.StartsWith("0")){ return false; }
            value = value.TrimEnd("1234567890".ToCharArray());
            return value.Equals(string.Empty);
        }

        public static string ToColumnName(int columnNumber)
        {
            var dividend = columnNumber;
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

        public static bool IsNumber(this string value)
        {
            double x;
            return double.TryParse(value, out x);
        }

        public static Drawing.Text ToDrawingText(this string str)
        {
            return new Drawing.Text(str);
        }

        public static Text ToText(this string str)
        {
            return new Text(str);
        }
        
        public static T Replace<T>(this T formula, string oldValue, string newValue) where T : OpenXmlLeafTextElement
        {
            formula.Text = formula.Text.Replace(oldValue, newValue);
            return formula;
        }

        public static T RegexReplace<T>(this T formula, string pattern, string replacement, RegexOptions options = RegexOptions.None) where T : OpenXmlLeafTextElement
        {
            Regex regEx = new Regex(pattern, options);
            formula.Text = regEx.Replace(formula.Text, replacement);
            return formula;
        }


        public static void RegexReplaceIn<T>(this OpenXmlElement element, string pattern, string replacement, RegexOptions options = RegexOptions.None) where T : OpenXmlLeafTextElement
        {
            Regex regEx = new Regex(pattern, options);
            var formulas = element.Descendants<T>();
            foreach (var formula in formulas)
            {
                formula.Text = regEx.Replace(formula.Text, replacement);
            }
        }

        public static void ReplaceIn<T>(this OpenXmlElement element, string oldValue, string newValue) where T : OpenXmlLeafTextElement
        {
            var formulas = element.Descendants<T>();
            foreach(var formula in formulas)
            {
                formula.Replace(oldValue, newValue);
            }
        }


    }
}
