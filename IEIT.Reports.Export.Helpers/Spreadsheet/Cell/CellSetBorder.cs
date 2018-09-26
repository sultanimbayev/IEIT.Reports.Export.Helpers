using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IEIT.Reports.Export.Helpers.Spreadsheet
{
    /// <summary>
    /// Управление границами ячейки
    /// </summary>
    public static class CellSetBorder
    {
        /// <summary>
        /// Задать стиль границы для ячейки
        /// </summary>
        /// <param name="cell">Объект ячейки</param>
        /// <param name="border">Объект границы</param>
        /// <returns></returns>
        public static Cell Set(this Cell cell, Border border)
        {
            var stylesheet = cell.GetWorkbookPart().GetStylesheet();
            var cellFormat = cell.StyleIndex != null ? stylesheet.GetCellFormat(cell.StyleIndex).CloneNode(true) as CellFormat : new CellFormat();
            var borderId = stylesheet.MakeBorder(border);
            cellFormat.BorderId = borderId;
            var newStyleId = stylesheet.MakeCellStyle(cellFormat);
            cell.StyleIndex = newStyleId;
            return cell;
        }

        /// <summary>
        /// Задать стиль границы для ячейки для определенной стороны
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="borderProp">
        ///     Свойство границы
        /// </param>
        /// <example>
        /// <code>
        ///     new LeftBorder();    
        ///     new RightBorder();    
        ///     new TopBorder();    
        ///     new BottomBorder();    
        /// </code>
        /// </example>
        /// <returns></returns>
        public static Cell Set(this Cell cell, BorderPropertiesType borderProp)
        {
            var stylesheet = cell.GetWorkbookPart().GetStylesheet();
            var cellFormat = cell.StyleIndex != null ? stylesheet.GetCellFormat(cell.StyleIndex).CloneNode(true) as CellFormat : new CellFormat();
            var borderId = cellFormat.BorderId;
            var border = borderId == null ? new Border() : stylesheet.GetBorder(borderId).CloneNode(true) as Border;

            if (borderProp is LeftBorder)
            {
                border.LeftBorder = borderProp as LeftBorder;
            }
            else if (borderProp is RightBorder)
            {
                border.RightBorder = borderProp as RightBorder;
            }
            else if (borderProp is BottomBorder)
            {
                border.BottomBorder = borderProp as BottomBorder;
            }
            else if (borderProp is TopBorder)
            {
                border.TopBorder = borderProp as TopBorder;
            }
            else
            {
                throw new NotImplementedException($"Невозможно установить стиль границы ячейки '{borderProp.GetType().Name}'");
            }

            var newBorderId = stylesheet.MakeBorder(border);
            cellFormat.BorderId = newBorderId;
            var newStyleId = stylesheet.MakeCellStyle(cellFormat);
            cell.StyleIndex = newStyleId;
            return cell;
        }
    }
}
