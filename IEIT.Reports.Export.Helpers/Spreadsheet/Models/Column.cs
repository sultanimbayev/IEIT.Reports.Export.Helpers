using x = DocumentFormat.OpenXml.Spreadsheet;
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
        public IEnumerable<x.Cell> Cells {
            get
            {
                return Worksheet.Descendants<x.Cell>().Where(c => c.CellReference != null && Utils.ToColumnName(c.CellReference.Value).Equals(ColumnName));
            }
        }

        /// <summary>
        /// Лист в котором находится этот столбец
        /// </summary>
        public x.Worksheet Worksheet { get; private set; }
        
        /// <summary>
        /// Конструктор столбца
        /// </summary>
        /// <param name="worksheet">Лист в котором находится столбец</param>
        /// <param name="address">Название столбца или его номер (начиная с 1-го)</param>
        public Column(x.Worksheet worksheet, string address)
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
        public Column(x.Worksheet worksheet, int columnNum)
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
        public Column(x.Worksheet worksheet, uint columnNum) : this(worksheet, (int)columnNum) { }

        /// <summary>
        /// Получить ячейку по номеру строки.
        /// Возвращает null если ячейки нет.
        /// </summary>
        /// <param name="rowNum">Номер строки</param>
        /// <returns>Ячейка с соответствующим адресом</returns>
        public x.Cell GetCell(uint rowNum)
        {
            return Cells.FirstOrDefault(c => Utils.ToRowNum(c.CellReference.Value) == rowNum);
        }

        /// <summary>
        /// Получить  ячейку на пересечении этой колонки с указанной строкой.
        /// Возвращает null если ячейки нет.
        /// </summary>
        /// <param name="rowNum">Номер строки</param>
        /// <returns>Ячейка с соответствующим адресом</returns>
        public x.Cell GetCell(int rowNum)
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
        public x.Cell MakeCell(uint rowNum)
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
        public x.Cell MakeCell(int rowNum)
        {
            return Worksheet.GetCell(ColumnName + rowNum);
        }

        /// <summary>
        /// Column width in characters
        /// </summary>
        /// <returns></returns>
        public double GetWidth()
        {
            var sheetFormatProps = Worksheet?.SheetFormatProperties;
            var defaultColWidth = sheetFormatProps?.DefaultColumnWidth;
            if (defaultColWidth == null)
            {
                defaultColWidth = 8.43;
            }
            var columns = Worksheet.Descendants<x.Column>();
            var columnProp = columns.Where(c => c.Min.HasValue && c.Min.Value <= ColumnNum && c.Max.HasValue && c.Max.Value >= ColumnNum).FirstOrDefault();
            return columnProp?.Width ?? defaultColWidth.Value;
        }


        /// <summary>
        /// Column width in pixels with 96dpi
        /// </summary>
        /// <returns>column width in pixels with 96dpi</returns>
        public double GetWidthInPixels()
        {
            var openXmlWidth = GetWidth();
            var widthInPixels = (openXmlWidth - 1) * 7d + 12;
            return widthInPixels;
        }

        /// <summary>
        /// Set column width in characters
        /// </summary>
        /// <param name="width">column width in characters</param>
        /// <returns></returns>
        public Column SetWidth(double width)
        {
            var columns = Worksheet.GetFirstChild<x.Columns>();
            if(columns == null)
            {
                columns = new x.Columns();
                Worksheet.Insert(columns).AfterOneOf(typeof(x.Dimension), typeof(x.SheetViews), typeof(x.SheetFormatProperties));
            }
            var columnProp = columns.Descendants<x.Column>().Where(c => c.Min.HasValue && c.Min.Value <= ColumnNum && c.Max.HasValue && c.Max.Value >= ColumnNum).FirstOrDefault();
            if(columnProp == null)
            {
                columnProp = new x.Column()
                {
                    Min = (uint)ColumnNum,
                    Max = (uint)ColumnNum
                };
                columns.Append(columnProp);
            }
            columnProp.CustomWidth = true;
            columnProp.Width = width;
            return this;
        }

        /// <summary>
        /// Set column width in pixels with 96dpi
        /// </summary>
        /// <param name="width">Width in pixels with 96dpi</param>
        /// <returns></returns>
        public Column SetWidthInPixels(double width)
        {
            var openXmlWidth = (width - 12) / 7d + 1;
            SetWidth(openXmlWidth);
            return this;
        }

    }
}
