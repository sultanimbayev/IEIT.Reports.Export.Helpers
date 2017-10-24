using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Drawing = DocumentFormat.OpenXml.Drawing;
using DrawingCharts = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet.Charts
{
    /// <summary>
    /// Класс с расширениями для работы с графиками
    /// </summary>
    public static class ChartHelper
    {
        /// <summary>
        /// Получить область с графиком(графиков) по названию ее верхней левой ячейки.
        /// </summary>
        /// <param name="ws">Лист в котором находится искомая область</param>
        /// <param name="cellAddress">Верхняя левая ячейка области графиков</param>
        /// <returns>Область графика(графиков) верхняя левая грань которой находится в заданной ячейке</returns>
        public static ChartPart ChartPartAt(this Worksheet ws, string cellAddress)
        {
            if(ws == null || ws.WorksheetPart == null) { return null; }

            var rowNum = Utils.ToRowNum(cellAddress);
            var colNum = Utils.ToColumNum(cellAddress);

            var rowId = (rowNum - 1).ToString();
            var colId = (colNum - 1).ToString();
            
            var anchors = ws.WorksheetPart.DrawingsPart.WorksheetDrawing.Elements<TwoCellAnchor>();

            var positionAnchor = anchors.FirstOrDefault(anc => anc.FromMarker.ColumnId.InnerText.Equals(colId) && anc.FromMarker.RowId.InnerText.Equals(rowId));

            if (positionAnchor == null)
            {
                return null;
            }

            var chartRef = positionAnchor.FirstDescendant<DrawingCharts.ChartReference>();

            if (chartRef == null || chartRef.Id == null || chartRef.Id.Value == null)
            {
                return null;
            }

            var chartPart = ws.WorksheetPart.DrawingsPart.GetPartById(chartRef.Id.Value) as ChartPart;

            return chartPart;

        }


        /// <summary>
        /// Перенести область графиков 
        /// </summary>
        /// <param name="chartPart">Область графиков для перемещения</param>
        /// <param name="topLeft">Левая верхняя позиция</param>
        /// <param name="bottomRight">Нижняя правая позиция</param>
        /// <returns>true - если операция выполнилась успешно, false - в обратном случае.</returns>
        public static bool RelocateTo(this ChartPart chartPart, Drawing.Spreadsheet.FromMarker topLeft, Drawing.Spreadsheet.ToMarker bottomRight)
        {
            var theAnchor = chartPart.Anchor();
            if(theAnchor == null) { return false; }

            theAnchor.FromMarker = topLeft;
            theAnchor.ToMarker = bottomRight;
            return true;
        }

        /// <summary>
        /// Перенести область графиков
        /// </summary>
        /// <param name="chartPart">Область графиков для перемещения</param>
        /// <param name="cellAddress">Позиция области графиков (ячейка, левый верхний край области)</param>
        /// <param name="columnOffset">Отклонение от левого края ячейки</param>
        /// <param name="rowOffset">Отклонение от верхнего края ячейки</param>
        /// <returns>true - если операция выполнилась успешно, false - в обратном случае.</returns>
        public static bool RelocateTo(this ChartPart chartPart, string cellAddress, long columnOffset = 0, long rowOffset = 0)
        {
            var columnNum = Utils.ToColumNum(cellAddress);
            var rowNum = Utils.ToRowNum(cellAddress);
            return chartPart.RelocateTo((int)columnNum, (int)rowNum, columnOffset, rowOffset);
        }

        /// <summary>
        /// Перенести область графиков
        /// </summary>
        /// <param name="chartPart">Область графиков для перемещения</param>
        /// <param name="columnNum">Номер колонки (соответствует левому краю области)</param>
        /// <param name="rowNum">Номер строки (соответствует верхнему краю области)</param>
        /// <param name="columnOffset">Отклонение от левого края ячейки</param>
        /// <param name="rowOffset">Отклонение от верхнего края ячейки</param>
        /// <returns>true - если операция выполнилась успешно, false - в обратном случае.</returns>
        public static bool RelocateTo(this ChartPart chartPart, int columnNum, int rowNum, long columnOffset = 0, long rowOffset = 0)
        {
            var theAnchor = chartPart.Anchor();
            if (theAnchor == null) { return false; }
            if (theAnchor.FromMarker == null || theAnchor.ToMarker == null) { return false; }

            int top, bottom;
            long topOffset, bottomOffset;
            int left, right;
            long leftOffset, rightOffset;

            try
            {
                top = int.Parse(theAnchor.FromMarker.RowId.InnerText);
                topOffset = long.Parse(theAnchor.FromMarker.RowOffset.InnerText);
                left = int.Parse(theAnchor.FromMarker.ColumnId.InnerText);
                leftOffset = long.Parse(theAnchor.FromMarker.ColumnOffset.InnerText);

                bottom = int.Parse(theAnchor.ToMarker.RowId.InnerText);
                bottomOffset = long.Parse(theAnchor.ToMarker.RowOffset.InnerText);
                right = int.Parse(theAnchor.ToMarker.ColumnId.InnerText);
                rightOffset = long.Parse(theAnchor.ToMarker.ColumnOffset.InnerText);
            }
            catch
            {
                return false;
            }

            var distanceX = right - left;
            var distanceXOffset = rightOffset - leftOffset;
            var distanceY = bottom - top;
            var distanceYOffset = bottomOffset - topOffset;

            int newTop = rowNum - 1;
            long newTopOffset = rowOffset;
            int newLeft = columnNum - 1;
            long newLeftOffset = columnOffset;

            int newBottom = newTop + distanceY;
            long newBottomOffset = newTopOffset + distanceYOffset;
            int newRight = newLeft + distanceX;
            long newRightOffset = newLeftOffset + distanceXOffset;

            var topLeft = new Drawing.Spreadsheet.FromMarker()
            {
                ColumnId = new ColumnId(newLeft.ToString()),
                ColumnOffset = new ColumnOffset(newLeftOffset.ToString()),
                RowId = new RowId(newTop.ToString()),
                RowOffset = new RowOffset(newTopOffset.ToString())
            };

            var bottomRight = new Drawing.Spreadsheet.ToMarker()
            {
                ColumnId = new ColumnId(newRight.ToString()),
                ColumnOffset = new ColumnOffset(newRightOffset.ToString()),
                RowId = new RowId(newBottom.ToString()),
                RowOffset = new RowOffset(newBottomOffset.ToString())
            };

            chartPart.RelocateTo(topLeft, bottomRight);
            return true;
        }

        /// <summary>
        /// Переместить область графиков
        /// </summary>
        /// <param name="chartPart">Часть документа с графиками</param>
        /// <param name="topLeft">Верхняя левая позиция области</param>
        /// <returns>true - если операция выполнилась успешно, false - в обратном случае.</returns>
        public static bool RelocateTo(this ChartPart chartPart, Drawing.Spreadsheet.FromMarker topLeft)
        {

            int newTop;
            long newTopOffset;
            int newLeft;
            long newLeftOffset;

            try
            {
                newTop = int.Parse(topLeft.RowId.InnerText);
                newTopOffset = long.Parse(topLeft.RowOffset.InnerText);
                newLeft = int.Parse(topLeft.ColumnId.InnerText);
                newLeftOffset = long.Parse(topLeft.ColumnOffset.InnerText);
            }
            catch
            {
                return false;
            }

            int columnNum = newLeft + 1;
            int rowNum = newTop + 1;
            return chartPart.RelocateTo(columnNum, rowNum, newLeftOffset, newTopOffset);
        }

        /// <summary>
        /// Получить информацию о местоположении области с графиками
        /// </summary>
        /// <param name="chartPart">Область информацию которой нужно получить</param>
        /// <returns>Сведения о положении области графиков на листе</returns>
        public static TwoCellAnchor Anchor(this ChartPart chartPart)
        {
            if (chartPart == null) { return null; }
            
            var drawingsPart = chartPart.GetParentParts().FirstOrDefault() as DrawingsPart;
            if(drawingsPart == null) { return null; }

            var partId = drawingsPart.GetIdOfPart(chartPart);
            var anchors = drawingsPart.WorksheetDrawing.Elements<TwoCellAnchor>();

            var theAnchor = anchors.FirstOrDefault(anc =>
            {
                var chartRef = anc.FirstDescendant<DrawingCharts.ChartReference>();
                if (chartRef == null) { return false; }
                return chartRef.Id.Value == partId;
            });

            return theAnchor;
        }

        /// <summary>
        /// Задать заголовок графику
        /// </summary>
        /// <param name="chart">график</param>
        /// <param name="newTitleStr">новый заголовок графика</param>
        /// <returns>true - при удачной замене заголовка, false - в обратном случае</returns>
        public static bool Title(this DrawingCharts.Chart chart, string newTitleStr)
        {

            if(chart == null)
            {
                return false;
            }

            var newTitle = new DrawingCharts.Title().From("Drawing\\ChartTitle");
            
            if(newTitle == null)
            {
                return false;
            }

            var run = newTitle.FirstDescendant<Drawing.Run>();

            if (run == null)
            {
                return false;
            }

            run.Text = newTitleStr.ToDrawingText();

            var oldTitle = chart.GetFirstChild<DrawingCharts.Title>();

            var pPr = oldTitle == null ? null : oldTitle.FirstDescendant<Drawing.ParagraphProperties>();
            var rPr = oldTitle == null ? null : oldTitle.FirstDescendant<Drawing.RunProperties>();

            var defPPr = newTitle.FirstDescendant<Drawing.ParagraphProperties>();
            var defRPr = newTitle.FirstDescendant<Drawing.RunProperties>();

            if(pPr != null && defPPr != null)
            {
                defPPr.Parent.ReplaceChild(pPr.CloneNode(true), defPPr);
            }

            if (rPr != null && defRPr != null)
            {
                defRPr.Parent.ReplaceChild(rPr.CloneNode(true), defRPr);
            }

            if(oldTitle == null)
            {
                chart.PrependChild(newTitle);
                return true;
            }

            var resultTitle = chart.ReplaceChild(newTitle, oldTitle);

            if (resultTitle.Equals(oldTitle))
            {
                return true;
            }

            return false;

        }

        /// <summary>
        /// Получить графики
        /// </summary>
        /// <param name="chartPart">Часть документа с графиками</param>
        /// <returns>Графики в выбранной части документа</returns>
        public static IEnumerable<DrawingCharts.Chart> Charts(this ChartPart chartPart)
        {
            return chartPart.ChartSpace.Charts();
        }

        /// <summary>
        /// Получить графики
        /// </summary>
        /// <param name="chartSpace">Область с графиками</param>
        /// <returns>Графики в выбранной области</returns>
        public static IEnumerable<DrawingCharts.Chart> Charts(this DrawingCharts.ChartSpace chartSpace)
        {
            return chartSpace.Descendants<DrawingCharts.Chart>();
        }

        /// <summary>
        /// Получить заголовок
        /// </summary>
        /// <param name="chart">График заголовок которогу нужно получить</param>
        /// <returns>Заголовок выбранного графика</returns>
        public static string Title(this DrawingCharts.Chart chart)
        {
            return chart.GetFirstChild<DrawingCharts.Title>().InnerText;
        }

        /// <summary>
        /// Задать формулу
        /// </summary>
        /// <param name="formula">Контейнер формулы</param>
        /// <param name="newFormula">Строка с новой формулой</param>
        /// <returns>Новая формула в виде строки</returns>
        public static string Text(this DrawingCharts.Formula formula, string newFormula)
        {
            return formula.Text = newFormula;
        }
        
    }
}
