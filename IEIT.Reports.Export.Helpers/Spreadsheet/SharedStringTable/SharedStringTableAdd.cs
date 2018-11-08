using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class SharedStringTableAdd
    {
        /// <summary>
        /// Добавить текст в таблицу <see cref="SharedStringTable"/>
        /// </summary>
        /// <param name="sst">Таблица с тектами</param>
        /// <param name="text">Добавляемый текст</param>
        /// <param name="rPr">Стиль добавляемого текста</param>
        /// <returns>Добавленыый элемент в <see cref="SharedStringTable"/> содержащий указанный текст</returns>
        public static SharedStringItem Add(this SharedStringTable sst, string text, RunProperties rPr = null)
        {
            var item = new SharedStringItem();
            if (rPr == null)
            {
                item.Text = new Text(text);
            }
            else
            {
                var run = new Run();
                run.Text = new Text(text);
                run.Append(rPr.CloneNode(true));
                item.Append(run);
            }
            sst.Append(item);
            if (sst.Count != null) { sst.Count.Value++; }
            if (sst.UniqueCount != null) { sst.UniqueCount.Value++; }
            return item;
        }
    }
}
