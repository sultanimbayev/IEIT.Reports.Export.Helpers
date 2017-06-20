using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Spreadsheet;
using IEIT.Reports.Export.Helpers.Spreadsheet.Intents;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers
{
    internal static class Common
    {
        /// <summary>
        /// Регулярное выражения соответствующее адресу ячейки
        /// </summary>
        internal const string RGX_PAT_CA = @"^[a-zA-Z]+\d+$";

        /// <summary>
        /// Регулярное выражение соответствующее ряду адресов ячеек
        /// </summary>
        internal const string RGX_PAT_CA_RANGE = @"^[a-zA-Z]+\d+:[a-zA-Z]+\d+$";


        /// <summary>
        /// Добавить текст в элемент типа <see cref="RstType"/> с указанным стилем
        /// </summary>
        /// <param name="item">Элемент к которому прибавляется текст</param>
        /// <param name="text">Добавляемый текст</param>
        /// <param name="rPr">Стиль добавляемого текста</param>
        public static void AppendText<T>(this T item, string text, RunProperties rPr = null) where T : RstType
        {
            if (item == null) { throw new ArgumentNullException($"item object to appending text is null"); }

            var lastElem = item.Elements<Run>().LastOrDefault();

            if (lastElem == null)
            {
                if (item.Text == null) { item.Text = new Text(); }
                if (rPr == null) { item.Text.Text += text; return; }
                var run2 = new Run();
                run2.Text = item.Text.CloneNode(true) as Text;
                item.Append(run2);
                item.Text.Remove();
                lastElem = run2;
            }

            if (lastElem == null || !lastElem.RunProperties.SameAs(rPr))
            {
                var run = new Run();
                run.RunProperties = rPr;
                run.Text = new Text(text);
                item.InsertAfter(run, lastElem);
                return;
            }

            if (lastElem.Text == null || string.IsNullOrEmpty(lastElem.Text.Text))
            {
                lastElem.Text = new Text(text);
                return;
            }

            lastElem.Text.Text += text;
            return;

        }

        /// <summary>
        /// Добавить элемент в дерево дочерних элементов. Возвращает "намерение" <see cref="InsertElementIntent{T}"/> для вставки элемента
        /// </summary>
        /// <typeparam name="T">Тип нового элемента</typeparam>
        /// <param name="parent">Родительский элемент в которую производится вставка</param>
        /// <param name="newChild">Вставляемый элемент</param>
        /// <returns>"Намерение" для вставки элемента</returns>
        public static InsertElementIntent<T> Insert<T>(this OpenXmlElement parent, T newChild) where T : OpenXmlElement
        {
            if (parent == null) { throw new ArgumentNullException("Parent object must be not null to insert child element!"); }
            return new InsertElementIntent<T>(parent, newChild);
        }

    }
}
