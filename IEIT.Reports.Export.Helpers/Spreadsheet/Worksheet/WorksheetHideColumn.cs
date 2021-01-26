using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetHideColumn
    {

        /// <summary>
        /// Скрытие колонки по ее названию.
        /// </summary>
        /// <param name="worksheet">Лист в котором будует скрыта колонка</param>
        /// <param name="columnName">Название колонки которая будет скрыта</param>
        /// <returns></returns>
        public static Worksheet HideColumn(this Worksheet worksheet, string columnName)
        {
            return HideColumn(worksheet, Utils.ToColumnNum(columnName));
        }

        /// <summary>
        /// Скрытие колонки по ее номеру.
        /// </summary>
        /// <param name="worksheet">Лист в котором будует скрыта колонка</param>
        /// <param name="columnIndx">Номер колонки которая будет скрыта</param>
        /// <returns></returns>
        public static Worksheet HideColumn(this Worksheet worksheet, int columnIndx)
        {
            if(columnIndx < 0) { return worksheet; }
            return HideColumn(worksheet, (uint)columnIndx);
        }

        /// <summary>
        /// Скрытие колонки по ее номеру.
        /// </summary>
        /// <param name="worksheet">Лист в котором будует скрыта колонка</param>
        /// <param name="columnIndx">Номер колонки которая будет скрыта</param>
        /// <returns></returns>
        public static Worksheet HideColumn(this Worksheet worksheet, uint columnIndx)
        {
            var columns = worksheet.GetFirstChild<Columns>();
            Column col = null;
            if (columns.Descendants<Column>().Any(c => c.Min == columnIndx))
            {
                col = columns.Descendants<Column>().Where(c => c.Min == columnIndx).First();
            }
            else
            {
                col = new Column();
                col.Min = columnIndx;
                col.Max = columnIndx;
                columns.Append(col);
            }
            col.Hidden = true;
            return worksheet;
        }
    }
}
