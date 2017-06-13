using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Drawing = DocumentFormat.OpenXml.Drawing;
using DrawingCharts = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;

namespace IEIT.Reports.Export.Helpers.Spreadsheet.Charts
{
    public static class ChartHelper
    {

        public static ChartPart GetChartPartByPosition(this Worksheet ws, string cellAddress)
        {
            if(ws == null || ws.WorksheetPart == null) { return null; }

            uint rowNum = Utils.ToRowNum(cellAddress);
            uint colNum = Utils.ToColumNum(cellAddress);

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


        public static bool SetTitle(this DrawingCharts.Chart chart, string newTitleStr)
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

        public static IEnumerable<DrawingCharts.Chart> GetCharts(this ChartPart chartPart)
        {
            return chartPart.ChartSpace.GetCharts();
        }

        public static IEnumerable<DrawingCharts.Chart> GetCharts(this DrawingCharts.ChartSpace chartSpace)
        {
            return chartSpace.Descendants<DrawingCharts.Chart>();
        }

        public static string GetTitle(this DrawingCharts.Chart chart)
        {
            return chart.GetFirstChild<DrawingCharts.Title>().InnerText;
        }


        public static string Set(this DrawingCharts.Formula formula, string newFormula)
        {
            return formula.Text = newFormula;
        }
        
    }
}
