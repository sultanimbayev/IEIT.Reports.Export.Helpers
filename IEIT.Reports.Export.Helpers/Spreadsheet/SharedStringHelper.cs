using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;
using System;
using System.Linq;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class SharedStringHelper
    {

        /// <summary>
        /// Получить <see cref="SharedStringItem"/> по его ID
        /// </summary>
        /// <param name="wbPart">Элемент <see cref="WorkbookPart"/></param>
        /// <param name="itemId">ID элемента <see cref="SharedStringItem"/></param>
        /// <returns>Элемент <see cref="SharedStringItem"/> с указанным ID</returns>
        internal static SharedStringItem GetSharedStringItem(this WorkbookPart wbPart, int itemId)
        {
            if (wbPart == null) { throw new ArgumentNullException("wbPart is null"); }
            return wbPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(itemId);
        }


        /// <summary>
        /// Получить <see cref="SharedStringItem"/> объект относящийся к данной ячейке.
        /// Возвращяет null если у данной ячейки нет такого объекта.
        /// </summary>
        /// <param name="cell">Ячейка документа</param>
        /// <returns> <see cref="SharedStringItem"/> объект относящийся к данной ячейке</returns>
        public static SharedStringItem GetSharedStringItem(this Cell cell)
        {
            if (cell == null) { throw new ArgumentNullException("Argument 'cell' must not be null!"); }
            if (cell.CellValue == null || cell.CellValue.Text == null) { return null; }
            if (cell.DataType != CellValues.SharedString) { return null; }
            var wbPart = cell.GetWorkbookPart();
            if (wbPart == null) { throw new InvalidDocumentStructureException("Given worksheet of given cell is not part of workbook!"); }
            var itemId = int.Parse(cell.CellValue.Text);
            return wbPart.GetSharedStringItem(itemId);
        }

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


        /// <summary>
        /// Переместить значение хранящиеся в данном объекте в SharedString.
        /// Не преобразует значения типа <see cref="CellValues.Boolean"/>
        /// <see cref="CellValues.Date"/> <see cref="CellValues.Error"/>
        /// <see cref="CellValues.Number"/> если не указан параметр <paramref name="force"/> как true
        /// </summary>
        /// <param name="cell">Ячейка, значение которой нужно сделать общим</param>
        /// <param name="force">Флаг "насильного" преобразования, если указан как true, то преобразует значение не смотря на его тип.
        /// А если указан false (по умолчанию), то преобразует только строковые значения.
        /// </param>
        /// <returns>Преобразованное значение <see cref="SharedStringItem"/> при удачном преобразовании, 
        /// либо null в обратном случае</returns>
        public static SharedStringItem MakeValueShared(this Cell cell, bool force = false)
        {

            if (cell.DataType != null && cell.DataType == CellValues.SharedString) { return cell.GetSharedStringItem(); }

            if (cell.DataType == null
                || cell.DataType == CellValues.Boolean
                || cell.DataType == CellValues.Date
                || cell.DataType == CellValues.Error
                || cell.DataType == CellValues.Number
                || force)
            {
                return null;
            }

            var wbPart = cell.GetWorkbookPart();
            if (wbPart == null) { throw new InvalidDocumentStructureException("Given worksheet of given cell is not part of workbook!"); }
            if (wbPart.SharedStringTablePart == null) { wbPart.AddNewPart<SharedStringTablePart>(); }
            if (wbPart.SharedStringTablePart.SharedStringTable == null) { wbPart.SharedStringTablePart.SharedStringTable = new SharedStringTable().From("SST.Empty"); }
            var sst = wbPart.SharedStringTablePart.SharedStringTable;
            if (cell.CellValue == null) { cell.CellValue = new CellValue(); }
            var itemIdx = sst.Elements().Count();
            SharedStringItem newItem;

            if (cell.DataType != null && cell.DataType == CellValues.InlineString)
            {
                var inStr = cell.InlineString;
                newItem = new SharedStringItem(inStr.Elements().Select(el => el.CloneNode(true)));
                sst.Append(newItem);
            }
            else
            {
                var text = cell.CellValue.Text;
                newItem = sst.Add(text);
            }

            cell.DataType = CellValues.SharedString;
            cell.CellValue.Text = itemIdx.ToString();
            cell.InlineString = null;
            return newItem;
        }

    }
}
