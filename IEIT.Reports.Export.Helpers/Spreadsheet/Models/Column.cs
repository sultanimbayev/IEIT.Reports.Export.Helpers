using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet.Models
{
    /// <summary>
    /// Класс для работы со столбцом в листе
    /// </summary>
    public class Column
    {
        /// <summary>
        /// Номер столбца (начиная с 1-го)
        /// </summary>
        public int ColumnNum { get; private set; }
        
        /// <summary>
        /// Название столбца, буквы латинского языка
        /// </summary>
        public string ColumnName { get; private set; }

        /// <summary>
        /// Ячейки хранящиеся в данном столбце
        /// </summary>
        public IEnumerable<Cell> Cells {
            get
            {
                return Worksheet.Descendants<Cell>().Where(c => c.CellReference != null && Utils.ToColumnName(c.CellReference.Value).Equals(ColumnName));
            }
        }

        /// <summary>
        /// Лист в котором находится этот столбец
        /// </summary>
        public Worksheet Worksheet { get; private set; }
        
        /// <summary>
        /// Конструктор столбца
        /// </summary>
        /// <param name="worksheet">Лист в котором находится столбец</param>
        /// <param name="address">Название столбца или его номер (начиная с 1-го)</param>
        public Column(Worksheet worksheet, string address)
        {
            Worksheet = worksheet;
            ColumnName = Utils.ToColumnName(address);
            ColumnNum = (int)Utils.ToColumNum(address);
        }


        /// <summary>
        /// Конструктор столбца
        /// </summary>
        /// <param name="worksheet">Лист в котором находится столбец</param>
        /// <param name="columnNum">номер столбца (начиная с 1-го)</param>
        public Column(Worksheet worksheet, int columnNum)
        {
            Worksheet = worksheet;
            ColumnName = Utils.ToColumnName(columnNum);
            ColumnNum = columnNum;
        }


        /// <summary>
        /// Конструктор столбца
        /// </summary>
        /// <param name="worksheet">Лист в котором находится столбец</param>
        /// <param name="columnNum">номер столбца (начиная с 1-го)</param>
        public Column(Worksheet worksheet, uint columnNum) : this(worksheet, (int)columnNum) { }

        /// <summary>
        /// Получить ячейку по номеру строки.
        /// Возвращает null если ячейки нет.
        /// </summary>
        /// <param name="rowNum">Номер строки</param>
        /// <returns>Ячейка с соответствующим адресом</returns>
        public Cell GetCell(uint rowNum)
        {
            return Cells.FirstOrDefault(c => Utils.ToRowNum(c.CellReference.Value) == rowNum);
        }

        /// <summary>
        /// Получить  ячейку на пересечении этой колонки с указанной строкой.
        /// Возвращает null если ячейки нет.
        /// </summary>
        /// <param name="rowNum">Номер строки</param>
        /// <returns>Ячейка с соответствующим адресом</returns>
        public Cell GetCell(int rowNum)
        {
            return GetCell((uint)rowNum);
        }

        /// <summary>
        /// Создать ячейку на пересечении этой колонки с указанной строкой если ее не существует.
        /// Если ячейка уже существует, тогда возвращяется существующая ячейка.
        /// Старая ячейка перезаписана не будет.
        /// Идентично методу <see cref="CellHelper.MakeCell(Worksheet, string)"/>
        /// </summary>
        /// <param name="rowNum">Номер строки</param>
        /// <returns>Ячейка с соответствующим адресом.</returns>
        public Cell MakeCell(uint rowNum)
        {
            return Worksheet.GetCell(ColumnName + rowNum);
        }

        /// <summary>
        /// Создать ячейку на пересечении этой колонки с указанной строкой если ее не существует.
        /// Если ячейка уже существует, тогда возвращяется существующая ячейка.
        /// Старая ячейка перезаписана не будет.
        /// Идентично методу <see cref="CellHelper.MakeCell(Worksheet, string)"/>
        /// </summary>
        /// <param name="rowNum">Номер строки</param>
        /// <returns>Ячейка с соответствующим адресом.</returns>
        public Cell MakeCell(int rowNum)
        {
            return Worksheet.GetCell(ColumnName + rowNum);
        }
    }
}
