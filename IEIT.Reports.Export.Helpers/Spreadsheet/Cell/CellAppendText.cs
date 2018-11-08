using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class CellAppendText
    {
        /// <summary>
        /// Добавление текста в ячейку
        /// </summary>
        /// <param name="cell">Ячейка в которую ведется запись</param>
        /// <param name="text">Добавляемый текст</param>
        /// <param name="styles">Стиль добавляемого текта</param>
        /// <returns>Всегда true</returns>
        public static bool AppendText(this Cell cell, string text, RunProperties styles = null)
        {
            TurnValueToInlineString(cell);
            if(cell.InlineString == null) { cell.InlineString = new InlineString(); }
            cell.InlineString.AppendText(text, styles);
            cell.CellValue = new CellValue();
            return true;
        }

        /// <summary>
        /// Переместить значение хранящиеся в данном объекте в форматированную строку
        /// </summary>
        /// <param name="cell"></param>
        public static void TurnValueToInlineString(Cell cell)
        {
            if (cell.DataType != null && cell.DataType == CellValues.InlineString) { return; }
            if (cell.CellValue == null) { cell.CellValue = new CellValue(); }
            InlineString newInStr;
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                var ssItem = cell.GetSharedStringItem();
                newInStr = new InlineString(ssItem.Elements().Select(el => el.CloneNode(true)));
            }
            else
            {
                var text = cell.CellValue.Text;
                newInStr = new InlineString();
                newInStr.Text = new Text(text);
            }

            cell.CellValue = new CellValue();
            cell.InlineString = newInStr;
            cell.DataType = CellValues.InlineString;
        }


    }
}
