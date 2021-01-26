using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Spreadsheet.Intents;
using IEIT.Reports.Export.Helpers.Styling;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetWriter
    {

        /// <summary>
        /// Записать значение в ячейку
        /// <para>- запись текста</para>
        /// <para>- запись числа</para>
        /// <para>- запись формулы</para>
        /// </summary>
        /// <param name="ws">Лист в который требуется записать значение</param>
        /// <param name="value">Значение которое нужно записать</param>
        /// <returns>Намерение, с помощью которого производится запись</returns>
        public static WriteIntent Write(this Worksheet ws, string value)
        {
            return new WriteIntent(ws).WithText(value);
        }

        /// <summary>
        /// Записать значение в ячейку
        /// </summary>
        /// <param name="ws">Лист в который требуется записать значение</param>
        /// <param name="value">Значение которое нужно записать</param>
        /// <returns>Намерение, с помощью которого производится запись</returns>
        public static WriteIntent Write(this Worksheet ws, object value)
        {
            var _val = value != null ? value.ToString() : "-";
            return new WriteIntent(ws).WithText(_val);
        }

        /// <summary>
        /// Записать текст в ячейку
        /// </summary>
        /// <param name="ws">Лист в который требуется записать значение</param>
        /// <param name="value">Значение которое нужно записать</param>
        /// <returns>Намерение, с помощью которого производится запись</returns>
        public static WriteIntent WriteText(this Worksheet ws, string text)
        {
            return new WriteIntent(ws, WriterActions._writeText).WithText(text);
        }

        /// <summary>
        /// Записать формулу в ячейку
        /// </summary>
        /// <param name="ws">Лист в который требуется записать значение</param>
        /// <param name="value">Значение которое нужно записать</param>
        /// <returns>Намерение, с помощью которого производится запись</returns>
        public static WriteIntent WriteFormula(this Worksheet ws, string formula)
        {
            return new WriteIntent(ws, WriterActions._writeFormula).WithText(formula);
        }


        /// <summary>
        /// Записать число в ячейку
        /// </summary>
        /// <param name="ws">Лист в который требуется записать значение</param>
        /// <param name="value">Значение которое нужно записать</param>
        /// <returns>Намерение, с помощью которого производится запись</returns>
        public static WriteIntent WriteNumber(this Worksheet ws, string number)
        {
            return new WriteIntent(ws, WriterActions._writeNumber).WithText(number);
        }

        /// <summary>
        /// Назначить стиль ячейки
        /// </summary>
        /// <param name="ws">Лист в которыом нужно изменить стиль</param>
        /// <param name="styleIndex">ID стиля</param>
        /// <returns>Намерение, с помощью которого производится запись</returns>
        public static WriteIntent SetStyle(this Worksheet ws, UInt32Value styleIndex)
        {
            return new WriteIntent(ws, WriterActions._writeAny).WithStyle(styleIndex);
        }

        /// <summary>
        /// Назначить стиль ячейки
        /// </summary>
        /// <param name="ws">Лист в которыом нужно изменить стиль</param>
        /// <param name="cellStyle">стиль ячейки</param>
        /// <returns>Намерение, с помощью которого производится запись</returns>
        public static WriteIntent SetStyle(this Worksheet ws, xlCellStyle cellStyle)
        {
            return new WriteIntent(ws, WriterActions._writeAny).WithStyle(cellStyle);
        }
    }
}
