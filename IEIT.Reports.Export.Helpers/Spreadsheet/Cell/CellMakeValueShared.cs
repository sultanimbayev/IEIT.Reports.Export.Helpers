using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;
using System.Linq;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class CellMakeValueShared
    {
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
            if (wbPart.SharedStringTablePart.SharedStringTable == null) { wbPart.SharedStringTablePart.SharedStringTable = new SharedStringTable() { Count = 0, UniqueCount = 0 }; }
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
