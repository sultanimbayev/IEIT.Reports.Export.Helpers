﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Exceptions;
using IEIT.Reports.Export.Helpers.Spreadsheet;
using System;
using System.Linq;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    public static class WorksheetHelper
    {

        /// <summary>
        /// Получить свойства листа
        /// </summary>
        /// <param name="worksheet">Объект листа</param>
        /// <returns>Объект содержащий свойства листа <see cref="Sheet"/></returns>
        internal static Sheet GetSheet(this Worksheet worksheet)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            if (worksheet.WorksheetPart == null) { throw new InvalidDocumentStructureException(); }
            var wbPart = worksheet.GetWorkbookPart();
            if (wbPart == null) { return null; }
            var wsPartId = wbPart.GetIdOfPart(worksheet.WorksheetPart);
            var sheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Id != null && s.Id.Value == wsPartId).FirstOrDefault();
            return sheet;
        }

        /// <summary>
        /// Получить часть документа который содержит все рабочие листы.
        /// </summary>
        /// <param name="worksheet">Рабочий лист</param>
        /// <returns>Часть документа который содержит все рабочие листы</returns>
        internal static WorkbookPart GetWorkbookPart(this Worksheet worksheet)
        {
            return worksheet.WorksheetPart.GetParentParts().FirstOrDefault() as WorkbookPart;
        }



        /// <summary>
        /// Получить объект ячейки. 
        /// Вызывает ошибку если ячейка еще не существует.
        /// <seealso cref="MakeCell(Worksheet, string)"/>
        /// </summary>
        /// <param name="worksheet">Лист в котором находится ячейка</param>
        /// <param name="cellAddress">Адрес ячейки</param>
        /// <returns>Объект ячейки который находится в данном листе по указанному адресу</returns>
        public static Cell GetCell(this Worksheet worksheet, string cellAddress)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }

            var rowNum = Utils.ToRowNum(cellAddress);
            var row = worksheet.GetRow(rowNum);

            var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == cellAddress);
            return cell;
        }

        /// <summary>
        /// Создать ячейку. Если ячейка уже создана в указанном месте, 
        /// тогда данный метод будет идентичен методу <see cref="GetCell(Worksheet, string)"/>
        /// </summary>
        /// <param name="worksheet">Лист в котором нужно создать ячейку</param>
        /// <param name="cellAddress">Адрес новой ячейки</param>
        /// <returns>Созданную ячейку, если ячейка не существовала</returns>
        public static Cell MakeCell(this Worksheet worksheet, string cellAddress)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            var rowNum = Utils.ToRowNum(cellAddress);
            var row = worksheet.MakeRow(rowNum);

            if (row == null) { throw new IncompleteActionException("Создание строки."); }

            var cell = row.Elements<Cell>()
                .Where(c => Utils.ToColumNum(c.CellReference.Value) >= Utils.ToColumNum(cellAddress))
                .OrderBy(c => Utils.ToColumNum(c.CellReference.Value))
                .FirstOrDefault();

            if (cell != null && cell.CellReference.Value.Equals(cellAddress, StringComparison.OrdinalIgnoreCase)) { return cell; }

            var newCell = new Cell();
            newCell.CellReference = cellAddress;
            row.InsertBefore(newCell, cell);

            return newCell;

        }

        /// <summary>
        /// Получить строку. 
        /// Вызывает ошибку если строка еще не существует. 
        /// <seealso cref="MakeRow(Worksheet, uint)"/>
        /// </summary>
        /// <param name="worksheet">Лист в котором находится требуемая строка</param>
        /// <param name="rowNum">Номер запрашиваемой строки</param>
        /// <returns>Объект строки</returns>
        public static Row GetRow(this Worksheet worksheet, uint rowNum)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            var wsData = worksheet.GetFirstChild<SheetData>();
            if (wsData == null) { throw new InvalidDocumentStructureException(); }
            var row = wsData.Elements<Row>().FirstOrDefault(r => r.RowIndex.Value == rowNum);
            return row;
        }


        /// <summary>
        /// Создать строку. Если строка уже существует, то данный метод будет
        /// идентичен методу <see cref="GetRow(Worksheet, uint)"/>
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="rowNum"></param>
        /// <returns></returns>
        public static Row MakeRow(this Worksheet worksheet, uint rowNum)
        {
            if (worksheet == null) { throw new ArgumentNullException("worksheet"); }
            var wsData = worksheet.GetFirstChild<SheetData>();
            if (wsData == null) { throw new InvalidDocumentStructureException(); }

            var row = wsData
                    .Elements<Row>()
                    .Where(r => r.RowIndex.Value >= rowNum)
                    .OrderBy(r => r.RowIndex.Value).FirstOrDefault();

            if (row != null && row.RowIndex == rowNum) { return row; }

            var newRow = new Row();
            newRow.RowIndex = rowNum;

            wsData.InsertBefore(newRow, row);

            return newRow;
        }


        /// <summary>
        /// Переименовать лист
        /// </summary>
        /// <param name="worksheet">Лист который ты хочешь переименовать</param>
        /// <param name="newName">Новое название листа</param>
        /// <param name="updateReferences">Заменить все ссылки к данному листу?</param>
        /// <returns>true при удачном переименовывании, false в обратном случае</returns>
        public static bool Rename(this Worksheet worksheet, string newName, bool updateReferences = true)
        {
            var sheet = worksheet.GetSheet();
            if (sheet == null) { return false; }

            var wbPart = worksheet.GetWorkbookPart();
            if (wbPart == null) { return false; }

            if (updateReferences)
            {
                var pattern = $"'?{sheet.Name}'?!";
                var replacement = $"'{newName}'!";

                var wsParts = wbPart.GetPartsOfType<WorksheetPart>();
                foreach (var wsPart in wsParts)
                {
                    //TODO: Улучшить обновление ссылок на лист
                    wsPart.RootElement.RegexReplaceIn<OpenXmlLeafTextElement>(pattern, replacement);
                    foreach (var chartPart in wsPart.DrawingsPart.ChartParts)
                    {
                        chartPart.RootElement.RegexReplaceIn<OpenXmlLeafTextElement>(pattern, replacement);
                    }
                }
            }

            return (sheet.Name = newName).Equals(newName);
        }


        /// <summary>
        /// Получить название листа
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public static string GetName(this Worksheet worksheet)
        {
            var sheet = worksheet.GetSheet();
            if (sheet == null) { return null; }
            return sheet.Name;
        }

    }
}