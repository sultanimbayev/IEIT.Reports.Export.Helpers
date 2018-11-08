using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet.Intents
{
    /// <summary>
    /// Класс для вставки копируемого значения
    /// </summary>
    public class PasteIntent
    {

        /// <summary>
        /// Лист из которого копируется значения ячеек
        /// </summary>
        private Worksheet srcWorksheet { get; set; }

        /// <summary>
        /// Область копируемых ячеек.
        /// </summary>
        private List<string> srcAddresses;

        /// <summary>
        /// Количество копируемых строк
        /// </summary>
        private uint _rowsCount;


        /// <summary>
        /// Количество копируемых колонок     
        /// </summary>
        private uint _colsCount;

        /// <summary>
        /// Конструктор для определения копируемых ячеек для последующей вставки
        /// </summary>
        /// <param name="sourceWorksheet">Лист где находятся копируемые ячейки</param>
        /// <param name="cellsRange">Область копируемых ячеек</param>
        public PasteIntent(Worksheet sourceWorksheet, string cellsRange)
        {
            srcWorksheet = sourceWorksheet;
            srcAddresses = Utils.CellAddressesFrom(cellsRange);
            var addrs = cellsRange.Split(':');
            if(addrs.Count() == 1) { _rowsCount = 1; _colsCount = 1; return; }

            _rowsCount = Utils.ToRowNum(addrs[1]) - Utils.ToRowNum(addrs[0]) + 1;
            _colsCount = Utils.ToColumNum(addrs[1]) - Utils.ToColumNum(addrs[0]) + 1;

        }


        /// <summary>
        /// Вставить элемент в тот же лист, в  указанный адрес ячейки
        /// </summary>
        /// <param name="targetCellAddr">Адрес ячейки для вставки, одна ячейка (верхняя-левая)</param>
        /// <returns>"Намерение для вставки копированного значения"</returns>
        public PasteIntent To(string targetCellAddr)
        {
            return To(srcWorksheet, targetCellAddr);
        }

        /// <summary>
        /// Вставить элемент в указанный лист и адрес ячейки
        /// </summary>
        /// <param name="targetWorksheet">Лист куда вставляется копированное значение</param>
        /// <param name="targetCellAddr">Адрес ячейки для вставки, одна ячейка (верхняя-левая)</param>
        /// <returns>"Намерение для вставки копированного значения"</returns>
        public PasteIntent To(Worksheet targetWorksheet, string targetCellAddr)
        {
            var trgColNum = Utils.ToColumNum(targetCellAddr);
            var trgRowNum = Utils.ToRowNum(targetCellAddr);

            var lastTrgCol = Utils.ToColumnName(trgColNum + _colsCount - 1);
            var lastTrgRow = trgRowNum + _rowsCount - 1;

            var lastTrgAddr = lastTrgCol + lastTrgRow.ToString();

            var trgAddresses = Utils.CellAddressesFrom($"{targetCellAddr}:{lastTrgAddr}");
            
            if(srcAddresses.Count() != trgAddresses.Count())
            {
                throw new Exception
                    (
                    @"Не удалось копировать ячейки из-за несоответсвия количества 
                    адресов копируемых и вставляемых ячеек. 
                    Проверьте метод Utils.CellAddressesFrom()"
                    );
            }

            var cnt = srcAddresses.Count();

            for (int i = 0; i < cnt; i++)
            {
                var srcCell = srcWorksheet.GetCell(srcAddresses[i]);
                if(srcCell == null) { continue; }
                srcCell.MakeValueShared();
                srcCell = srcCell.CloneNode(true) as Cell;
                var trgCell = targetWorksheet.GetCell(trgAddresses[i]);
                srcCell.CellReference = trgCell.CellReference;
                trgCell.ReplaceBy(srcCell);
            }

            return this;
        }

    }
}
